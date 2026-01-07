#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
KiCad VIAステッチプラグイン（最適化版 V2.2.0）
このプラグインは、KiCadのPCBエディタ(Pcbnew)で選択された銅エリアにVIAを自動配置します。
熱伝導や電流伝導を改善するためのVIAステッチングを効率的に行うことができます。
KiCad 9.0.2対応版 - 空間インデックス化とプログレスバー対応

V2.2.0 修正点:
- トラック（配線）との衝突回避を追加
- VIA間クリアランスの数値入力対応（DRCルール代替）
- トラッククリアランスの数値入力対応

V2.1.0 修正点:
- PADに被っている位置へのVIA配置を回避
- 基板外にはみ出した位置へのVIA配置を回避
- 禁止領域（Keepout Zone）へのVIA配置を回避
- ゾーンの正確な形状（ポリゴン）を使用した配置判定
"""

import os
import sys
import json
import math
import pcbnew
import wx
import wx.grid
import random
import traceback
from datetime import datetime

# プラグインのバージョン
PLUGIN_VERSION = "2.2.0"

# デフォルト設定
DEFAULT_SETTINGS = {
    "via_size": 0.6,        # VIAサイズ (mm)
    "drill_size": 0.3,      # ドリルサイズ (mm)
    "h_spacing": 1.27,      # 水平方向の間隔 (mm)
    "v_spacing": 1.27,      # 垂直方向の間隔 (mm)
    "edge_clearance": 0.5,  # エッジクリアランス (mm)
    "h_offset": 0.0,        # 水平方向のオフセット (mm)
    "v_offset": 0.0,        # 垂直方向のオフセット (mm)
    "pattern": "grid",      # 配置パターン (grid, boundary, spiral)
    "randomize": False,     # ランダム配置
    "clear_plugin_vias": True,  # プラグインで配置したVIAのみをクリア
    "group_name": "ViaStitching",  # VIAグループ名
    "pad_clearance": 0.2,   # パッドからのクリアランス (mm)
    "check_keepout": True,  # 禁止領域をチェック
    "check_board_outline": True,  # 基板外形をチェック
    # V2.2.0 新規追加
    "via_clearance": 0.25,  # VIA間クリアランス (mm)
    "track_clearance": 0.2, # トラックからのクリアランス (mm)
    "check_tracks": True,   # トラックとの衝突をチェック
}


class SpatialIndex:
    """空間インデックス - VIAの高速検索用"""
    
    def __init__(self, grid_size):
        """
        Args:
            grid_size: グリッドセルのサイズ（nm単位）
        """
        self.grid_size = grid_size
        self.grid = {}  # {(grid_x, grid_y): [(pos, via), ...]}
    
    def _get_grid_coords(self, pos):
        """位置からグリッド座標を取得"""
        return (pos.x // self.grid_size, pos.y // self.grid_size)
    
    def add_via(self, pos, via):
        """VIAを空間インデックスに追加"""
        grid_coords = self._get_grid_coords(pos)
        if grid_coords not in self.grid:
            self.grid[grid_coords] = []
        self.grid[grid_coords].append((pos, via))
    
    def get_nearby_vias(self, pos, radius):
        """指定位置周辺のVIAを高速検索"""
        grid_coords = self._get_grid_coords(pos)
        grid_x, grid_y = grid_coords
        
        # 検索範囲を計算（周辺のグリッドセルも含める）
        grid_radius = int(math.ceil(radius / self.grid_size))
        
        nearby_vias = []
        for dx in range(-grid_radius, grid_radius + 1):
            for dy in range(-grid_radius, grid_radius + 1):
                check_coords = (grid_x + dx, grid_y + dy)
                if check_coords in self.grid:
                    for via_pos, via in self.grid[check_coords]:
                        # 実際の距離をチェック
                        dx_real = pos.x - via_pos.x
                        dy_real = pos.y - via_pos.y
                        distance = math.sqrt(dx_real * dx_real + dy_real * dy_real)
                        if distance <= radius:
                            nearby_vias.append((via_pos, via, distance))
        
        return nearby_vias


class PadSpatialIndex:
    """パッド用空間インデックス - パッドとの衝突判定を高速化"""
    
    def __init__(self, grid_size):
        self.grid_size = grid_size
        self.grid = {}  # {(grid_x, grid_y): [(center, pad_info), ...]}
    
    def _get_grid_coords(self, pos):
        return (pos.x // self.grid_size, pos.y // self.grid_size)
    
    def add_pad(self, center, pad_info):
        """パッドを空間インデックスに追加
        pad_info: dict with 'pad', 'bbox', 'radius' (円形近似半径)
        """
        grid_coords = self._get_grid_coords(center)
        if grid_coords not in self.grid:
            self.grid[grid_coords] = []
        self.grid[grid_coords].append((center, pad_info))
    
    def get_nearby_pads(self, pos, radius):
        """指定位置周辺のパッドを高速検索"""
        grid_coords = self._get_grid_coords(pos)
        grid_x, grid_y = grid_coords
        
        grid_radius = int(math.ceil(radius / self.grid_size)) + 1
        
        nearby_pads = []
        for dx in range(-grid_radius, grid_radius + 1):
            for dy in range(-grid_radius, grid_radius + 1):
                check_coords = (grid_x + dx, grid_y + dy)
                if check_coords in self.grid:
                    for pad_center, pad_info in self.grid[check_coords]:
                        dx_real = pos.x - pad_center.x
                        dy_real = pos.y - pad_center.y
                        distance = math.sqrt(dx_real * dx_real + dy_real * dy_real)
                        # パッドの半径も考慮
                        if distance <= radius + pad_info['radius']:
                            nearby_pads.append((pad_center, pad_info, distance))
        
        return nearby_pads


class TrackSpatialIndex:
    """トラック用空間インデックス - トラックとの衝突判定を高速化"""
    
    def __init__(self, grid_size):
        self.grid_size = grid_size
        self.grid = {}  # {(grid_x, grid_y): [track_info, ...]}
    
    def _get_grid_coords(self, pos):
        return (int(pos.x) // self.grid_size, int(pos.y) // self.grid_size)
    
    def _get_segment_grid_cells(self, start, end):
        """セグメントが通過するすべてのグリッドセルを取得"""
        cells = set()
        
        # セグメントのバウンディングボックス
        min_x = min(start.x, end.x)
        max_x = max(start.x, end.x)
        min_y = min(start.y, end.y)
        max_y = max(start.y, end.y)
        
        # グリッドセル範囲
        start_gx = int(min_x) // self.grid_size
        end_gx = int(max_x) // self.grid_size
        start_gy = int(min_y) // self.grid_size
        end_gy = int(max_y) // self.grid_size
        
        # すべての通過セルを追加
        for gx in range(start_gx, end_gx + 1):
            for gy in range(start_gy, end_gy + 1):
                cells.add((gx, gy))
        
        return cells
    
    def add_track(self, track, track_info):
        """トラックを空間インデックスに追加
        track_info: dict with 'track', 'start', 'end', 'width', 'net_code'
        """
        cells = self._get_segment_grid_cells(track_info['start'], track_info['end'])
        
        for cell in cells:
            if cell not in self.grid:
                self.grid[cell] = []
            self.grid[cell].append(track_info)
    
    def get_nearby_tracks(self, pos, radius):
        """指定位置周辺のトラックを高速検索"""
        grid_coords = self._get_grid_coords(pos)
        grid_x, grid_y = grid_coords
        
        grid_radius = int(math.ceil(radius / self.grid_size)) + 1
        
        nearby_tracks = set()  # 重複排除用
        result = []
        
        for dx in range(-grid_radius, grid_radius + 1):
            for dy in range(-grid_radius, grid_radius + 1):
                check_coords = (grid_x + dx, grid_y + dy)
                if check_coords in self.grid:
                    for track_info in self.grid[check_coords]:
                        # 同じトラックを重複追加しない
                        track_id = id(track_info['track'])
                        if track_id not in nearby_tracks:
                            nearby_tracks.add(track_id)
                            result.append(track_info)
        
        return result


def point_to_segment_distance(point, seg_start, seg_end):
    """点とセグメント（線分）の最短距離を計算
    
    Args:
        point: チェックする点 (VECTOR2I)
        seg_start: セグメントの始点 (VECTOR2I)
        seg_end: セグメントの終点 (VECTOR2I)
    
    Returns:
        float: 最短距離
    """
    px, py = float(point.x), float(point.y)
    x1, y1 = float(seg_start.x), float(seg_start.y)
    x2, y2 = float(seg_end.x), float(seg_end.y)
    
    # セグメントの長さの2乗
    dx = x2 - x1
    dy = y2 - y1
    seg_len_sq = dx * dx + dy * dy
    
    if seg_len_sq == 0:
        # 始点と終点が同じ（点）
        return math.sqrt((px - x1) ** 2 + (py - y1) ** 2)
    
    # 点からセグメントへの射影位置（0-1の範囲にクランプ）
    t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / seg_len_sq))
    
    # セグメント上の最近接点
    proj_x = x1 + t * dx
    proj_y = y1 + t * dy
    
    # 距離を計算
    return math.sqrt((px - proj_x) ** 2 + (py - proj_y) ** 2)


class BoardGeometryChecker:
    """基板形状チェッカー - 基板外形、禁止領域、パッド、トラックなどをチェック"""
    
    def __init__(self, board, via_size, pad_clearance, track_clearance, via_clearance):
        self.board = board
        self.via_size = via_size
        self.via_radius = via_size / 2
        self.pad_clearance = pad_clearance
        self.track_clearance = track_clearance
        self.via_clearance = via_clearance
        
        # パッドの空間インデックスを構築
        self.pad_index = PadSpatialIndex(int(via_size * 3))
        self._build_pad_index()
        
        # トラックの空間インデックスを構築
        self.track_index = TrackSpatialIndex(int(via_size * 3))
        self._build_track_index()
        
        # 基板外形のポリゴンを取得
        self.board_outline = self._get_board_outline()
        
        # 禁止領域（Keepout Zone）のリストを取得
        self.keepout_zones = self._get_keepout_zones()
        
        print(f"パッド数: {self.pad_count}")
        print(f"トラック数: {self.track_count}")
        print(f"禁止領域数: {len(self.keepout_zones)}")
        print(f"基板外形取得: {'成功' if self.board_outline else '失敗'}")
    
    def _build_pad_index(self):
        """パッドの空間インデックスを構築"""
        self.pad_count = 0
        
        for footprint in self.board.GetFootprints():
            for pad in footprint.Pads():
                try:
                    center = pad.GetPosition()
                    bbox = pad.GetBoundingBox()
                    
                    # パッドサイズから円形近似半径を計算
                    pad_size = pad.GetSize()
                    radius = max(pad_size.x, pad_size.y) / 2
                    
                    pad_info = {
                        'pad': pad,
                        'bbox': bbox,
                        'radius': radius,
                        'net_code': pad.GetNetCode()
                    }
                    
                    self.pad_index.add_pad(center, pad_info)
                    self.pad_count += 1
                except Exception as e:
                    print(f"パッド情報取得エラー: {e}")
    
    def _build_track_index(self):
        """トラックの空間インデックスを構築"""
        self.track_count = 0
        
        for item in self.board.GetTracks():
            # VIAは除外（トラックのみ）
            if item.Type() == pcbnew.PCB_TRACE_T:
                try:
                    start = item.GetStart()
                    end = item.GetEnd()
                    width = item.GetWidth()
                    net_code = item.GetNetCode()
                    
                    track_info = {
                        'track': item,
                        'start': start,
                        'end': end,
                        'width': width,
                        'net_code': net_code
                    }
                    
                    self.track_index.add_track(item, track_info)
                    self.track_count += 1
                except Exception as e:
                    print(f"トラック情報取得エラー: {e}")
    
    def _get_board_outline(self):
        """基板外形（Edge.Cuts）からポリゴンを取得"""
        try:
            # 基板の境界を取得
            bbox = self.board.GetBoardEdgesBoundingBox()
            if bbox.GetWidth() == 0 or bbox.GetHeight() == 0:
                print("基板外形が見つかりません（BoundingBox使用）")
                return None
            
            # Edge.Cutsレイヤーの描画要素を収集
            edge_segments = []
            for drawing in self.board.GetDrawings():
                if drawing.GetLayer() == pcbnew.Edge_Cuts:
                    edge_segments.append(drawing)
            
            if not edge_segments:
                print("Edge.Cutsレイヤーに描画要素がありません")
                # BoundingBoxを代替として使用
                return {
                    'type': 'bbox',
                    'bbox': bbox
                }
            
            # セグメントからポリゴンを構築を試みる
            # 簡略化: BoundingBoxを使用し、Edge.Cuts要素の存在を確認
            return {
                'type': 'segments',
                'segments': edge_segments,
                'bbox': bbox
            }
            
        except Exception as e:
            print(f"基板外形取得エラー: {e}")
            traceback.print_exc()
            return None
    
    def _get_keepout_zones(self):
        """禁止領域（VIA配置禁止のKeepout Zone）を取得"""
        keepout_zones = []
        
        try:
            for zone in self.board.Zones():
                try:
                    # KiCad 9.x API
                    if hasattr(zone, 'GetIsRuleArea'):
                        is_rule_area = zone.GetIsRuleArea()
                    else:
                        is_rule_area = zone.IsRuleArea() if hasattr(zone, 'IsRuleArea') else False
                    
                    if is_rule_area:
                        # VIA配置禁止かどうかチェック
                        do_not_allow_vias = False
                        if hasattr(zone, 'GetDoNotAllowVias'):
                            do_not_allow_vias = zone.GetDoNotAllowVias()
                        elif hasattr(zone, 'DoNotAllowVias'):
                            do_not_allow_vias = zone.DoNotAllowVias()
                        
                        if do_not_allow_vias:
                            keepout_zones.append(zone)
                            print(f"禁止領域検出: VIA配置禁止")
                except Exception as e:
                    print(f"ゾーンチェックエラー: {e}")
                    continue
        except Exception as e:
            print(f"禁止領域取得エラー: {e}")
        
        return keepout_zones
    
    def is_point_in_board(self, point):
        """点が基板内にあるかチェック"""
        if self.board_outline is None:
            return True  # 外形がない場合は常にTrue
        
        try:
            bbox = self.board_outline['bbox']
            margin = self.via_radius
            
            # BoundingBoxによる簡易チェック
            if (point.x < bbox.GetLeft() + margin or 
                point.x > bbox.GetRight() - margin or
                point.y < bbox.GetTop() + margin or 
                point.y > bbox.GetBottom() - margin):
                return False
            
            # より詳細なチェック（Edge.Cutsセグメントがある場合）
            if self.board_outline['type'] == 'segments':
                # セグメントベースの判定（オプション）
                # ここでは簡略化のためBoundingBoxのみ使用
                pass
            
            return True
            
        except Exception as e:
            print(f"基板内判定エラー: {e}")
            return True  # エラー時は配置を許可
    
    def is_point_in_keepout(self, point):
        """点が禁止領域内にあるかチェック"""
        for zone in self.keepout_zones:
            try:
                # KiCad APIでゾーン内判定
                # SHAPE_POLY_SETを使用
                outline = zone.Outline()
                if outline:
                    # VECTOR2Iに変換
                    pt = pcbnew.VECTOR2I(int(point.x), int(point.y))
                    if outline.Contains(pt):
                        return True
                else:
                    # フォールバック: BoundingBoxでチェック
                    bbox = zone.GetBoundingBox()
                    if (bbox.GetLeft() <= point.x <= bbox.GetRight() and
                        bbox.GetTop() <= point.y <= bbox.GetBottom()):
                        return True
            except Exception as e:
                # フォールバック: BoundingBoxでチェック
                try:
                    bbox = zone.GetBoundingBox()
                    if (bbox.GetLeft() <= point.x <= bbox.GetRight() and
                        bbox.GetTop() <= point.y <= bbox.GetBottom()):
                        return True
                except:
                    pass
        
        return False
    
    def is_point_on_pad(self, point, via_net_code=None):
        """点がパッド上にあるかチェック
        
        Args:
            point: チェックする位置
            via_net_code: VIAのネットコード（同じネットのパッドは許可する場合に使用）
        
        Returns:
            bool: パッド上にある場合True
        """
        # 検索範囲 = VIA半径 + パッドクリアランス + 余裕
        search_radius = self.via_radius + self.pad_clearance
        
        nearby_pads = self.pad_index.get_nearby_pads(point, search_radius)
        
        for pad_center, pad_info, distance in nearby_pads:
            # 同じネットのパッドは許可（オプション）
            # if via_net_code is not None and pad_info['net_code'] == via_net_code:
            #     continue
            
            # より正確なチェック: パッド半径 + VIA半径 + クリアランス
            min_distance = pad_info['radius'] + self.via_radius + self.pad_clearance
            if distance < min_distance:
                return True
        
        return False
    
    def is_point_on_track(self, point, via_net_code=None):
        """点がトラック上または近接しているかチェック
        
        Args:
            point: チェックする位置
            via_net_code: VIAのネットコード（同じネットのトラックは許可）
        
        Returns:
            bool: トラック上または近接している場合True
        """
        # 検索範囲 = VIA半径 + トラッククリアランス + 最大トラック幅の余裕
        max_track_width = 1000000  # 1mm相当（nm単位）、実際はトラックごとに判定
        search_radius = self.via_radius + self.track_clearance + max_track_width
        
        nearby_tracks = self.track_index.get_nearby_tracks(point, search_radius)
        
        for track_info in nearby_tracks:
            # 同じネットのトラックは許可
            if via_net_code is not None and track_info['net_code'] == via_net_code:
                continue
            
            # 点とトラック（セグメント）の距離を計算
            distance = point_to_segment_distance(
                point, 
                track_info['start'], 
                track_info['end']
            )
            
            # 必要クリアランス = VIA半径 + トラック幅/2 + クリアランス
            min_distance = self.via_radius + track_info['width'] / 2 + self.track_clearance
            
            if distance < min_distance:
                return True
        
        return False
    
    def can_place_via(self, point, via_net_code=None, check_board=True, check_keepout=True, 
                      check_pads=True, check_tracks=True):
        """VIAを配置できるかの総合チェック
        
        Returns:
            tuple: (can_place: bool, reason: str)
        """
        # 基板外形チェック
        if check_board and not self.is_point_in_board(point):
            return False, "基板外"
        
        # 禁止領域チェック
        if check_keepout and self.is_point_in_keepout(point):
            return False, "禁止領域内"
        
        # パッド衝突チェック
        if check_pads and self.is_point_on_pad(point, via_net_code):
            return False, "パッド上"
        
        # トラック衝突チェック
        if check_tracks and self.is_point_on_track(point, via_net_code):
            return False, "トラック上"
        
        return True, "OK"


def get_settings_path():
    """設定ファイルのパスを取得"""
    kicad_config_path = os.path.join(os.path.expanduser("~"), ".config", "kicad")
    if not os.path.exists(kicad_config_path):
        os.makedirs(kicad_config_path)
    return os.path.join(kicad_config_path, "viastitching_settings.json")


def load_settings():
    """設定をファイルから読み込み"""
    settings_path = get_settings_path()
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r') as f:
                loaded = json.load(f)
                # デフォルト設定とマージ（新しい設定項目に対応）
                merged = DEFAULT_SETTINGS.copy()
                merged.update(loaded)
                return merged
        except Exception as e:
            print(f"設定ファイルの読み込みエラー: {e}")
    return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    """設定をファイルに保存"""
    settings_path = get_settings_path()
    try:
        with open(settings_path, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"設定ファイルの保存エラー: {e}")


class ViaStitchingDialog(wx.Dialog):
    """VIAステッチングのダイアログ"""
    
    def __init__(self, parent, board, selected_area, settings):
        wx.Dialog.__init__(self, parent, title="VIAステッチング v" + PLUGIN_VERSION, 
                          style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        
        self.board = board
        self.selected_area = selected_area
        self.settings = settings
        self.net_info = None
        self.net_data = []
        self.pattern_data = [
            {"name": "格子状配置", "value": "grid"},
            {"name": "境界配置", "value": "boundary"},
            {"name": "スパイラル配置", "value": "spiral"}
        ]
        
        # ネット情報の取得
        if selected_area is not None:
            try:
                self.net_info = selected_area.GetNetCode()
                print(f"選択エリアのネットコード: {self.net_info}")
            except Exception as e:
                print(f"ネット情報取得エラー: {e}")
                self.net_info = 0
        
        # UIの構築
        self.build_ui()
        
        # 初期値の設定
        self.set_initial_values()

        # アクションの設定
        self.action_radio.SetSelection(0)
        self.clear_plugin_vias_checkbox.Enable(False)

        # レイアウトの調整
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.Layout()
        self.Fit()
        
        # イベントハンドラの設定
        self.Bind(wx.EVT_BUTTON, self.on_ok, id=wx.ID_OK)
        self.Bind(wx.EVT_BUTTON, self.on_cancel, id=wx.ID_CANCEL)
        self.Bind(wx.EVT_BUTTON, self.on_clear, id=wx.ID_CLEAR)
        self.Bind(wx.EVT_RADIOBOX, self.on_action_changed, self.action_radio)
        self.Bind(wx.EVT_CHECKBOX, self.on_randomize_changed, self.randomize_checkbox)
        self.Bind(wx.EVT_CHOICE, self.on_pattern_changed, self.pattern_choice)
        
        # プレビューの更新
        self.update_preview()

    def build_ui(self):
        """UIの構築"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # ネット選択
        net_sizer = wx.BoxSizer(wx.HORIZONTAL)
        net_label = wx.StaticText(self, wx.ID_ANY, "ネット名:")
        self.net_choice = wx.Choice(self, wx.ID_ANY)
        net_sizer.Add(net_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        net_sizer.Add(self.net_choice, 1, wx.ALL|wx.EXPAND, 5)
        main_sizer.Add(net_sizer, 0, wx.EXPAND, 5)
        
        # パラメータ設定
        param_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "VIAパラメータ"), wx.VERTICAL)
        
        # VIAサイズとドリルサイズ
        size_sizer = wx.BoxSizer(wx.HORIZONTAL)
        via_size_label = wx.StaticText(self, wx.ID_ANY, "VIAサイズ (mm):")
        self.via_size_text = wx.TextCtrl(self, wx.ID_ANY)
        drill_size_label = wx.StaticText(self, wx.ID_ANY, "ドリルサイズ (mm):")
        self.drill_size_text = wx.TextCtrl(self, wx.ID_ANY)
        size_sizer.Add(via_size_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        size_sizer.Add(self.via_size_text, 1, wx.ALL|wx.EXPAND, 5)
        size_sizer.Add(drill_size_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        size_sizer.Add(self.drill_size_text, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(size_sizer, 0, wx.EXPAND, 5)
        
        # 間隔設定
        spacing_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h_spacing_label = wx.StaticText(self, wx.ID_ANY, "水平間隔 (mm):")
        self.h_spacing_text = wx.TextCtrl(self, wx.ID_ANY)
        v_spacing_label = wx.StaticText(self, wx.ID_ANY, "垂直間隔 (mm):")
        self.v_spacing_text = wx.TextCtrl(self, wx.ID_ANY)
        spacing_sizer.Add(h_spacing_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        spacing_sizer.Add(self.h_spacing_text, 1, wx.ALL|wx.EXPAND, 5)
        spacing_sizer.Add(v_spacing_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        spacing_sizer.Add(self.v_spacing_text, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(spacing_sizer, 0, wx.EXPAND, 5)
        
        # オフセット設定
        offset_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h_offset_label = wx.StaticText(self, wx.ID_ANY, "水平オフセット (mm):")
        self.h_offset_text = wx.TextCtrl(self, wx.ID_ANY)
        v_offset_label = wx.StaticText(self, wx.ID_ANY, "垂直オフセット (mm):")
        self.v_offset_text = wx.TextCtrl(self, wx.ID_ANY)
        offset_sizer.Add(h_offset_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        offset_sizer.Add(self.h_offset_text, 1, wx.ALL|wx.EXPAND, 5)
        offset_sizer.Add(v_offset_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        offset_sizer.Add(self.v_offset_text, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(offset_sizer, 0, wx.EXPAND, 5)
        
        # エッジクリアランス
        edge_clearance_sizer = wx.BoxSizer(wx.HORIZONTAL)
        edge_clearance_label = wx.StaticText(self, wx.ID_ANY, "エッジクリアランス (mm):")
        self.clearance_text = wx.TextCtrl(self, wx.ID_ANY)
        edge_clearance_sizer.Add(edge_clearance_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        edge_clearance_sizer.Add(self.clearance_text, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(edge_clearance_sizer, 0, wx.EXPAND, 5)
        
        # パターン選択
        pattern_sizer = wx.BoxSizer(wx.HORIZONTAL)
        pattern_label = wx.StaticText(self, wx.ID_ANY, "配置パターン:")
        self.pattern_choice = wx.Choice(self, wx.ID_ANY)
        for pattern in self.pattern_data:
            self.pattern_choice.Append(pattern["name"])
        pattern_sizer.Add(pattern_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        pattern_sizer.Add(self.pattern_choice, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(pattern_sizer, 0, wx.EXPAND, 5)
        
        # ランダム化オプション
        self.randomize_checkbox = wx.CheckBox(self, wx.ID_ANY, "ランダム配置")
        param_sizer.Add(self.randomize_checkbox, 0, wx.ALL, 5)
        
        main_sizer.Add(param_sizer, 0, wx.ALL|wx.EXPAND, 5)
        
        # クリアランス設定（V2.2.0で拡張）
        clearance_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "クリアランス設定"), wx.VERTICAL)
        
        # パッドクリアランス
        pad_clearance_sizer = wx.BoxSizer(wx.HORIZONTAL)
        pad_clearance_label = wx.StaticText(self, wx.ID_ANY, "パッドクリアランス (mm):")
        self.pad_clearance_text = wx.TextCtrl(self, wx.ID_ANY)
        pad_clearance_sizer.Add(pad_clearance_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        pad_clearance_sizer.Add(self.pad_clearance_text, 1, wx.ALL|wx.EXPAND, 5)
        clearance_sizer.Add(pad_clearance_sizer, 0, wx.EXPAND, 5)
        
        # トラッククリアランス（V2.2.0新規）
        track_clearance_sizer = wx.BoxSizer(wx.HORIZONTAL)
        track_clearance_label = wx.StaticText(self, wx.ID_ANY, "トラッククリアランス (mm):")
        self.track_clearance_text = wx.TextCtrl(self, wx.ID_ANY)
        track_clearance_sizer.Add(track_clearance_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        track_clearance_sizer.Add(self.track_clearance_text, 1, wx.ALL|wx.EXPAND, 5)
        clearance_sizer.Add(track_clearance_sizer, 0, wx.EXPAND, 5)
        
        # VIA間クリアランス（V2.2.0新規）
        via_clearance_sizer = wx.BoxSizer(wx.HORIZONTAL)
        via_clearance_label = wx.StaticText(self, wx.ID_ANY, "VIA間クリアランス (mm):")
        self.via_clearance_text = wx.TextCtrl(self, wx.ID_ANY)
        via_clearance_sizer.Add(via_clearance_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        via_clearance_sizer.Add(self.via_clearance_text, 1, wx.ALL|wx.EXPAND, 5)
        clearance_sizer.Add(via_clearance_sizer, 0, wx.EXPAND, 5)
        
        main_sizer.Add(clearance_sizer, 0, wx.ALL|wx.EXPAND, 5)
        
        # 衝突回避オプション（V2.2.0で拡張）
        collision_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "衝突回避設定"), wx.VERTICAL)
        self.check_pads_checkbox = wx.CheckBox(self, wx.ID_ANY, "パッドとの衝突を回避")
        self.check_pads_checkbox.SetValue(True)
        self.check_tracks_checkbox = wx.CheckBox(self, wx.ID_ANY, "トラック（配線）との衝突を回避")  # V2.2.0新規
        self.check_tracks_checkbox.SetValue(True)
        self.check_keepout_checkbox = wx.CheckBox(self, wx.ID_ANY, "禁止領域（Keepout）を回避")
        self.check_keepout_checkbox.SetValue(True)
        self.check_board_outline_checkbox = wx.CheckBox(self, wx.ID_ANY, "基板外形を考慮")
        self.check_board_outline_checkbox.SetValue(True)
        collision_sizer.Add(self.check_pads_checkbox, 0, wx.ALL, 5)
        collision_sizer.Add(self.check_tracks_checkbox, 0, wx.ALL, 5)  # V2.2.0新規
        collision_sizer.Add(self.check_keepout_checkbox, 0, wx.ALL, 5)
        collision_sizer.Add(self.check_board_outline_checkbox, 0, wx.ALL, 5)
        main_sizer.Add(collision_sizer, 0, wx.ALL|wx.EXPAND, 5)
        
        # プレビューエリア
        preview_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "プレビュー"), wx.VERTICAL)
        self.preview_panel = wx.Panel(self, wx.ID_ANY, size=(300, 200))
        self.preview_panel.SetBackgroundColour(wx.WHITE)
        self.preview_panel.Bind(wx.EVT_PAINT, self.on_paint_preview)
        preview_sizer.Add(self.preview_panel, 1, wx.ALL|wx.EXPAND, 5)
        main_sizer.Add(preview_sizer, 1, wx.ALL|wx.EXPAND, 5)
        
        # アクション選択
        action_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "アクション"), wx.VERTICAL)
        action_choices = ["配置", "クリア"]
        self.action_radio = wx.RadioBox(self, wx.ID_ANY, "アクション", choices=action_choices, majorDimension=1)
        main_sizer.Add(self.action_radio, 0, wx.ALL|wx.EXPAND, 5)

        # プラグインVIAのみクリアオプション
        self.clear_plugin_vias_checkbox = wx.CheckBox(self, wx.ID_ANY, "プラグインで配置したVIAのみをクリア")
        action_sizer.Add(self.clear_plugin_vias_checkbox, 0, wx.ALL, 5)
        main_sizer.Add(action_sizer, 0, wx.ALL|wx.EXPAND, 5)
        
        # ボタン
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.ok_button = wx.Button(self, wx.ID_OK, "OK")
        self.cancel_button = wx.Button(self, wx.ID_CANCEL, "キャンセル")
        self.clear_button = wx.Button(self, wx.ID_CLEAR, "リセット")
        button_sizer.Add(self.ok_button, 0, wx.ALL, 5)
        button_sizer.Add(self.cancel_button, 0, wx.ALL, 5)
        button_sizer.Add(self.clear_button, 0, wx.ALL, 5)
        main_sizer.Add(button_sizer, 0, wx.ALIGN_RIGHT, 5)
        
        self.SetSizer(main_sizer)
    
    def set_initial_values(self):
        """初期値の設定"""
        self.populate_net_list()
        
        self.via_size_text.SetValue(str(self.settings["via_size"]))
        self.drill_size_text.SetValue(str(self.settings["drill_size"]))
        self.h_spacing_text.SetValue(str(self.settings["h_spacing"]))
        self.v_spacing_text.SetValue(str(self.settings["v_spacing"]))
        self.h_offset_text.SetValue(str(self.settings["h_offset"]))
        self.v_offset_text.SetValue(str(self.settings["v_offset"]))
        self.clearance_text.SetValue(str(self.settings["edge_clearance"]))
        self.pad_clearance_text.SetValue(str(self.settings.get("pad_clearance", 0.2)))
        self.track_clearance_text.SetValue(str(self.settings.get("track_clearance", 0.2)))  # V2.2.0
        self.via_clearance_text.SetValue(str(self.settings.get("via_clearance", 0.25)))    # V2.2.0
        
        pattern_index = 0
        for i, pattern in enumerate(self.pattern_data):
            if pattern["value"] == self.settings["pattern"]:
                pattern_index = i
                break
        self.pattern_choice.SetSelection(pattern_index)
        
        self.randomize_checkbox.SetValue(self.settings["randomize"])
        self.clear_plugin_vias_checkbox.SetValue(self.settings["clear_plugin_vias"])
        self.check_keepout_checkbox.SetValue(self.settings.get("check_keepout", True))
        self.check_board_outline_checkbox.SetValue(self.settings.get("check_board_outline", True))
        self.check_tracks_checkbox.SetValue(self.settings.get("check_tracks", True))  # V2.2.0
        
        self.action_radio.SetSelection(0)
    
    def populate_net_list(self):
        """ネットリストの設定"""
        self.net_choice.Clear()
        self.net_data = []
        
        try:
            nets = self.board.GetNetInfo().NetsByName()
            
            for net_name, net_item in nets.items():
                try:
                    net_id = net_item.GetNetCode()
                except:
                    try:
                        net_id = net_item.GetNet()
                    except:
                        try:
                            net_id = net_item.GetCode()
                        except:
                            net_id = 0
                
                self.net_data.append({"name": str(net_name), "id": net_id})
            
            self.net_data.sort(key=lambda x: x["name"])
            
            selected_index = 0
            for i, net in enumerate(self.net_data):
                self.net_choice.Append(net["name"])
                if net["id"] == self.net_info:
                    selected_index = i
            
            if self.net_choice.GetCount() > 0:
                self.net_choice.SetSelection(selected_index)
        except Exception as e:
            print(f"ネットリスト取得エラー: {e}")
            traceback.print_exc()
            self.net_data.append({"name": "未接続", "id": 0})
            self.net_choice.Append("未接続")
            self.net_choice.SetSelection(0)
    
    def update_preview(self):
        """プレビューの更新"""
        self.preview_panel.Refresh()
    
    def on_paint_preview(self, event):
        """プレビューの描画"""
        dc = wx.PaintDC(self.preview_panel)
        dc.Clear()
        
        width, height = self.preview_panel.GetSize()
        
        pattern_index = self.pattern_choice.GetSelection()
        if pattern_index != wx.NOT_FOUND:
            pattern = self.pattern_data[pattern_index]["value"]
        else:
            pattern = "grid"
        
        try:
            h_spacing = float(self.h_spacing_text.GetValue())
            v_spacing = float(self.v_spacing_text.GetValue())
            h_offset = float(self.h_offset_text.GetValue())
            v_offset = float(self.v_offset_text.GetValue())
            clearance = float(self.clearance_text.GetValue())
            randomize = self.randomize_checkbox.GetValue()
        except ValueError:
            return
        
        scale = min(width, height) / 20.0
        draw_width = width - 20
        draw_height = height - 20
        
        dc.SetBrush(wx.Brush(wx.Colour(200, 200, 200)))
        dc.SetPen(wx.Pen(wx.Colour(100, 100, 100), 2))
        dc.DrawRectangle(10, 10, draw_width, draw_height)
        
        dc.SetBrush(wx.Brush(wx.Colour(255, 255, 0)))
        dc.SetPen(wx.Pen(wx.Colour(0, 0, 0), 1))
        
        if pattern == "grid":
            rows = int(draw_height / (v_spacing * scale))
            cols = int(draw_width / (h_spacing * scale))
            
            for row in range(rows):
                for col in range(cols):
                    x = 10 + col * h_spacing * scale + h_offset * scale
                    y = 10 + row * v_spacing * scale + v_offset * scale
                    
                    if (x >= 10 + clearance * scale and 
                        x <= 10 + draw_width - clearance * scale and 
                        y >= 10 + clearance * scale and 
                        y <= 10 + draw_height - clearance * scale):
                        
                        if randomize:
                            x += random.uniform(-h_spacing * scale * 0.2, h_spacing * scale * 0.2)
                            y += random.uniform(-v_spacing * scale * 0.2, v_spacing * scale * 0.2)
                        
                        dc.DrawCircle(int(x), int(y), 3)
        
        elif pattern == "boundary":
            perimeter = 2 * (draw_width + draw_height)
            spacing = h_spacing * scale
            num_vias = int(perimeter / spacing)
            
            for i in range(num_vias):
                pos = i * spacing
                
                if pos < draw_width:
                    x = 10 + pos
                    y = 10 + clearance * scale
                elif pos < draw_width + draw_height:
                    x = 10 + draw_width - clearance * scale
                    y = 10 + (pos - draw_width)
                elif pos < 2 * draw_width + draw_height:
                    x = 10 + draw_width - (pos - draw_width - draw_height)
                    y = 10 + draw_height - clearance * scale
                else:
                    x = 10 + clearance * scale
                    y = 10 + draw_height - (pos - 2 * draw_width - draw_height)
                
                if randomize:
                    if pos < draw_width or (pos >= 2 * draw_width + draw_height):
                        x += random.uniform(-spacing * 0.1, spacing * 0.1)
                        y += random.uniform(0, spacing * 0.2)
                    else:
                        x -= random.uniform(0, spacing * 0.2)
                        y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                dc.DrawCircle(int(x), int(y), 3)
        
        elif pattern == "spiral":
            center_x = 10 + draw_width / 2
            center_y = 10 + draw_height / 2
            max_radius = min(draw_width, draw_height) / 2 - clearance * scale
            
            spacing = h_spacing * scale
            theta = 0
            radius = clearance * scale
            
            while radius <= max_radius:
                x = center_x + radius * math.cos(theta)
                y = center_y + radius * math.sin(theta)
                
                if randomize:
                    x += random.uniform(-spacing * 0.1, spacing * 0.1)
                    y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                dc.DrawCircle(int(x), int(y), 3)
                
                theta += spacing / radius
                radius = clearance * scale + spacing * theta / (2 * math.pi)
    
    def on_action_changed(self, event):
        """アクション変更時の処理"""
        if self.action_radio.GetSelection() == 1:
            self.clear_plugin_vias_checkbox.Enable(True)
        else:
            self.clear_plugin_vias_checkbox.Enable(False)
        self.Layout()
        self.update_preview()
    
    def on_randomize_changed(self, event):
        """ランダム化オプション変更時の処理"""
        self.update_preview()
    
    def on_pattern_changed(self, event):
        """パターン変更時の処理"""
        self.update_preview()
    
    def on_ok(self, event):
        """OKボタン押下時の処理"""
        try:
            self.settings["via_size"] = float(self.via_size_text.GetValue())
            self.settings["drill_size"] = float(self.drill_size_text.GetValue())
            self.settings["h_spacing"] = float(self.h_spacing_text.GetValue())
            self.settings["v_spacing"] = float(self.v_spacing_text.GetValue())
            self.settings["h_offset"] = float(self.h_offset_text.GetValue())
            self.settings["v_offset"] = float(self.v_offset_text.GetValue())
            self.settings["edge_clearance"] = float(self.clearance_text.GetValue())
            self.settings["pad_clearance"] = float(self.pad_clearance_text.GetValue())
            self.settings["track_clearance"] = float(self.track_clearance_text.GetValue())  # V2.2.0
            self.settings["via_clearance"] = float(self.via_clearance_text.GetValue())      # V2.2.0
            
            pattern_index = self.pattern_choice.GetSelection()
            if pattern_index != wx.NOT_FOUND:
                self.settings["pattern"] = self.pattern_data[pattern_index]["value"]
            
            self.settings["randomize"] = self.randomize_checkbox.GetValue()
            self.settings["clear_plugin_vias"] = self.clear_plugin_vias_checkbox.GetValue()
            self.settings["check_keepout"] = self.check_keepout_checkbox.GetValue()
            self.settings["check_board_outline"] = self.check_board_outline_checkbox.GetValue()
            self.settings["check_pads"] = self.check_pads_checkbox.GetValue()
            self.settings["check_tracks"] = self.check_tracks_checkbox.GetValue()  # V2.2.0
            
            save_settings(self.settings)
            self.EndModal(wx.ID_OK)
        except ValueError as e:
            wx.MessageBox(f"入力値が不正です: {e}", "エラー", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"エラーが発生しました: {e}", "エラー", wx.OK | wx.ICON_ERROR)
    
    def on_cancel(self, event):
        """キャンセルボタン押下時の処理"""
        self.EndModal(wx.ID_CANCEL)
    
    def on_clear(self, event):
        """クリアボタン押下時の処理"""
        self.via_size_text.SetValue(str(DEFAULT_SETTINGS["via_size"]))
        self.drill_size_text.SetValue(str(DEFAULT_SETTINGS["drill_size"]))
        self.h_spacing_text.SetValue(str(DEFAULT_SETTINGS["h_spacing"]))
        self.v_spacing_text.SetValue(str(DEFAULT_SETTINGS["v_spacing"]))
        self.h_offset_text.SetValue(str(DEFAULT_SETTINGS["h_offset"]))
        self.v_offset_text.SetValue(str(DEFAULT_SETTINGS["v_offset"]))
        self.clearance_text.SetValue(str(DEFAULT_SETTINGS["edge_clearance"]))
        self.pad_clearance_text.SetValue(str(DEFAULT_SETTINGS["pad_clearance"]))
        self.track_clearance_text.SetValue(str(DEFAULT_SETTINGS["track_clearance"]))  # V2.2.0
        self.via_clearance_text.SetValue(str(DEFAULT_SETTINGS["via_clearance"]))      # V2.2.0
        
        for i, pattern in enumerate(self.pattern_data):
            if pattern["value"] == DEFAULT_SETTINGS["pattern"]:
                self.pattern_choice.SetSelection(i)
                break
        
        self.randomize_checkbox.SetValue(DEFAULT_SETTINGS["randomize"])
        self.clear_plugin_vias_checkbox.SetValue(DEFAULT_SETTINGS["clear_plugin_vias"])
        self.check_keepout_checkbox.SetValue(DEFAULT_SETTINGS["check_keepout"])
        self.check_board_outline_checkbox.SetValue(DEFAULT_SETTINGS["check_board_outline"])
        self.check_pads_checkbox.SetValue(True)
        self.check_tracks_checkbox.SetValue(DEFAULT_SETTINGS["check_tracks"])  # V2.2.0
        
        self.update_preview()
    
    def get_action(self):
        """選択されたアクションを取得"""
        if self.action_radio.GetSelection() == 0:
            return "fill"
        else:
            return "clear"
    
    def get_net_code(self):
        """選択されたネットコードを取得"""
        try:
            index = self.net_choice.GetSelection()
            if index != wx.NOT_FOUND and index < len(self.net_data):
                return self.net_data[index]["id"]
            return 0
        except Exception as e:
            print(f"ネットコード取得エラー: {e}")
            return 0


class ViaStitchingPlugin(pcbnew.ActionPlugin):
    """VIAステッチングプラグイン"""
    
    def defaults(self):
        """プラグインのデフォルト設定"""
        self.name = "VIAステッチングV2"
        self.category = "配置"
        self.description = "選択された銅エリアにVIAを自動配置します（衝突回避機能付き）"
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), "via_stitcher_v2.png")
    
    def Run(self):
        """プラグインの実行"""
        try:
            board = pcbnew.GetBoard()
            print("ボード取得成功")
            
            selected_zones = []
            
            try:
                for zone in board.Zones():
                    if zone.IsSelected():
                        # 禁止領域（ルールエリア）は除外
                        is_rule_area = False
                        if hasattr(zone, 'GetIsRuleArea'):
                            is_rule_area = zone.GetIsRuleArea()
                        elif hasattr(zone, 'IsRuleArea'):
                            is_rule_area = zone.IsRuleArea()
                        
                        if not is_rule_area:
                            selected_zones.append(zone)
                            print(f"選択されたゾーン: {zone.GetNetname()}")
            except Exception as e:
                print(f"Zones()メソッドエラー: {e}")
                try:
                    selection = board.GetSelection()
                    for item in selection:
                        if item.IsType(pcbnew.PCB_ZONE_AREA_T):
                            selected_zones.append(item)
                except Exception as e:
                    print(f"GetSelection()メソッドエラー: {e}")
                    try:
                        for i in range(board.GetAreaCount()):
                            zone = board.GetArea(i)
                            if zone.IsSelected():
                                selected_zones.append(zone)
                    except Exception as e:
                        print(f"GetArea()メソッドエラー: {e}")
            
            if not selected_zones:
                wx.MessageBox("銅エリア（ゾーン）を選択してください\n※禁止領域（ルールエリア）は選択できません", 
                             "エラー", wx.OK | wx.ICON_ERROR)
                return
            
            print(f"選択されたゾーン数: {len(selected_zones)}")
            
            settings = load_settings()
            
            with ViaStitchingDialog(None, board, selected_zones[0], settings) as dlg:
                if dlg.ShowModal() == wx.ID_OK:
                    try:
                        action = dlg.get_action()
                        net_code = dlg.get_net_code()
                        print(f"アクション: {action}, ネットコード: {net_code}")
                        
                        if net_code is None:
                            wx.MessageBox("有効なネットが選択されていません", "エラー", wx.OK | wx.ICON_ERROR)
                            return
                        
                        if action == "fill":
                            self.fill_zones_with_vias_optimized(board, selected_zones, net_code, settings)
                        else:
                            self.clear_vias(board, selected_zones, net_code, settings)
                    except Exception as e:
                        traceback.print_exc()
                        wx.MessageBox(f"処理中にエラーが発生しました: {e}", "エラー", wx.OK | wx.ICON_ERROR)
        
        except Exception as e:
            traceback.print_exc()
            wx.MessageBox(f"エラーが発生しました: {e}", "エラー", wx.OK | wx.ICON_ERROR)
    
    def fill_zones_with_vias_optimized(self, board, zones, net_code, settings):
        """最適化されたVIA配置（衝突回避機能付き）"""
        # パラメータの取得
        via_size = pcbnew.FromMM(settings["via_size"])
        drill_size = pcbnew.FromMM(settings["drill_size"])
        h_spacing = pcbnew.FromMM(settings["h_spacing"])
        v_spacing = pcbnew.FromMM(settings["v_spacing"])
        h_offset = pcbnew.FromMM(settings["h_offset"])
        v_offset = pcbnew.FromMM(settings["v_offset"])
        edge_clearance = pcbnew.FromMM(settings["edge_clearance"])
        pad_clearance = pcbnew.FromMM(settings.get("pad_clearance", 0.2))
        track_clearance = pcbnew.FromMM(settings.get("track_clearance", 0.2))  # V2.2.0
        via_clearance = pcbnew.FromMM(settings.get("via_clearance", 0.25))     # V2.2.0
        pattern = settings["pattern"]
        randomize = settings["randomize"]
        group_name = settings["group_name"]
        
        # 衝突回避オプション
        check_pads = settings.get("check_pads", True)
        check_keepout = settings.get("check_keepout", True)
        check_board_outline = settings.get("check_board_outline", True)
        check_tracks = settings.get("check_tracks", True)  # V2.2.0
        
        # 空間インデックス初期化
        spatial_index = SpatialIndex(int(via_size * 1.5))
        
        # 基板形状チェッカーの初期化（V2.2.0でトラッククリアランス追加）
        geometry_checker = BoardGeometryChecker(
            board, via_size, pad_clearance, track_clearance, via_clearance
        )
        
        # 既存のVIAを空間インデックスに登録
        existing_vias = []
        for item in board.GetTracks():
            if item.Type() == pcbnew.PCB_VIA_T:
                pos = item.GetPosition()
                spatial_index.add_via(pos, item)
                existing_vias.append(item)
        
        print(f"既存VIA数: {len(existing_vias)}")
        
        # 全候補位置を事前計算
        all_candidate_positions = []
        for zone in zones:
            positions = self.calculate_candidate_positions(
                zone, pattern, h_spacing, v_spacing, 
                h_offset, v_offset, edge_clearance, randomize
            )
            all_candidate_positions.extend(positions)
        
        total_candidates = len(all_candidate_positions)
        print(f"総候補位置数: {total_candidates}")
        
        if total_candidates == 0:
            wx.MessageBox("配置可能な位置がありません", "情報", wx.OK | wx.ICON_INFORMATION)
            return
        
        # プログレスダイアログの作成
        progress_dlg = wx.ProgressDialog(
            "VIA配置中...",
            "候補位置を検証しています...",
            maximum=total_candidates,
            parent=None,
            style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE | wx.PD_CAN_ABORT | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME
        )
        
        # 統計情報（V2.2.0でトラック追加）
        skip_reasons = {
            "基板外": 0,
            "禁止領域内": 0,
            "パッド上": 0,
            "トラック上": 0,  # V2.2.0新規
            "VIA近接": 0,
        }
        
        try:
            valid_positions = []
            processed = 0
            
            for pos in all_candidate_positions:
                if processed % 10 == 0:
                    cont, skip = progress_dlg.Update(processed, f"位置検証中... ({processed}/{total_candidates})")
                    if not cont:
                        wx.MessageBox("処理がキャンセルされました", "情報", wx.OK | wx.ICON_INFORMATION)
                        return
                
                # 衝突チェック（V2.2.0でトラックチェック追加）
                can_place, reason = geometry_checker.can_place_via(
                    pos, net_code,
                    check_board=check_board_outline,
                    check_keepout=check_keepout,
                    check_pads=check_pads,
                    check_tracks=check_tracks  # V2.2.0
                )
                
                if not can_place:
                    if reason in skip_reasons:
                        skip_reasons[reason] += 1
                    processed += 1
                    continue
                
                # VIA間の距離チェック（V2.2.0でvia_clearance使用）
                if not self.check_drc_fast(pos, via_size, via_clearance, spatial_index):
                    skip_reasons["VIA近接"] += 1
                    processed += 1
                    continue
                
                valid_positions.append(pos)
                processed += 1
            
            valid_count = len(valid_positions)
            print(f"有効位置数: {valid_count}")
            print(f"スキップ理由: {skip_reasons}")
            
            if valid_count == 0:
                skip_info = "\n".join([f"  {k}: {v}" for k, v in skip_reasons.items() if v > 0])
                wx.MessageBox(f"配置可能な位置がありません\n\nスキップ理由:\n{skip_info}", 
                             "情報", wx.OK | wx.ICON_INFORMATION)
                return
            
            progress_dlg.Update(0, f"VIA作成中... (0/{valid_count})")
            
            # グループの作成
            group = None
            try:
                group = pcbnew.PCB_GROUP(board)
                group.SetName(f"{group_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
                board.Add(group)
                print("グループ作成成功")
            except Exception as e:
                print(f"グループ作成エラー: {e}")
            
            # VIAをバッチで作成
            created_vias = 0
            for i, pos in enumerate(valid_positions):
                if i % 5 == 0:
                    cont, skip = progress_dlg.Update(i, f"VIA作成中... ({i}/{valid_count})")
                    if not cont:
                        wx.MessageBox(f"{created_vias}個のVIAを配置して処理を中断しました", 
                                     "情報", wx.OK | wx.ICON_INFORMATION)
                        break
                
                via = self.add_via(board, pos, net_code, via_size, drill_size)
                if via:
                    spatial_index.add_via(pos, via)
                    
                    if group:
                        try:
                            group.AddItem(via)
                        except:
                            try:
                                group.AddPcbItem(via)
                            except:
                                pass
                    
                    created_vias += 1
            
            # 結果の表示
            skip_info = "\n".join([f"  {k}: {v}" for k, v in skip_reasons.items() if v > 0])
            wx.MessageBox(
                f"{created_vias}個のVIAを配置しました\n\n"
                f"スキップされた位置:\n{skip_info if skip_info else '  なし'}", 
                "完了", wx.OK | wx.ICON_INFORMATION
            )
            
        finally:
            progress_dlg.Destroy()
        
        pcbnew.Refresh()
    
    def calculate_candidate_positions(self, zone, pattern, h_spacing, v_spacing, 
                                      h_offset, v_offset, edge_clearance, randomize):
        """候補位置を事前計算（ゾーンのポリゴン形状を考慮）"""
        positions = []
        bbox = zone.GetBoundingBox()
        
        # ゾーンのアウトラインを取得
        zone_outline = None
        try:
            zone_outline = zone.Outline()
        except:
            pass
        
        if pattern == "grid":
            x_start = bbox.GetLeft() + edge_clearance + h_offset
            y_start = bbox.GetTop() + edge_clearance + v_offset
            x_end = bbox.GetRight() - edge_clearance
            y_end = bbox.GetBottom() - edge_clearance
            
            x = x_start
            while x <= x_end:
                y = y_start
                while y <= y_end:
                    if randomize:
                        rand_x = x + random.uniform(-h_spacing * 0.2, h_spacing * 0.2)
                        rand_y = y + random.uniform(-v_spacing * 0.2, v_spacing * 0.2)
                        pos = pcbnew.VECTOR2I(int(rand_x), int(rand_y))
                    else:
                        pos = pcbnew.VECTOR2I(int(x), int(y))
                    
                    # ゾーンの実際の形状でチェック
                    if self.is_point_in_zone_polygon(pos, zone, zone_outline, edge_clearance):
                        positions.append(pos)
                    
                    y += v_spacing
                x += h_spacing
        
        elif pattern == "boundary":
            perimeter = 2 * (bbox.GetWidth() + bbox.GetHeight())
            spacing = h_spacing
            num_vias = int(perimeter / spacing)
            
            for i in range(num_vias):
                t = i / num_vias
                perimeter_pos = t * perimeter
                
                if perimeter_pos < bbox.GetWidth():
                    x = bbox.GetLeft() + perimeter_pos
                    y = bbox.GetTop() + edge_clearance
                elif perimeter_pos < bbox.GetWidth() + bbox.GetHeight():
                    x = bbox.GetRight() - edge_clearance
                    y = bbox.GetTop() + (perimeter_pos - bbox.GetWidth())
                elif perimeter_pos < 2 * bbox.GetWidth() + bbox.GetHeight():
                    x = bbox.GetRight() - (perimeter_pos - bbox.GetWidth() - bbox.GetHeight())
                    y = bbox.GetBottom() - edge_clearance
                else:
                    x = bbox.GetLeft() + edge_clearance
                    y = bbox.GetBottom() - (perimeter_pos - 2 * bbox.GetWidth() - bbox.GetHeight())
                
                if randomize:
                    rand_dist = random.uniform(0, spacing * 0.2)
                    if perimeter_pos < bbox.GetWidth() or (perimeter_pos >= 2 * bbox.GetWidth() + bbox.GetHeight()):
                        x += random.uniform(-spacing * 0.1, spacing * 0.1)
                        y += rand_dist
                    else:
                        x -= rand_dist
                        y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                pos = pcbnew.VECTOR2I(int(x), int(y))
                
                if self.is_point_in_zone_polygon(pos, zone, zone_outline, edge_clearance):
                    positions.append(pos)
        
        elif pattern == "spiral":
            center_x = (bbox.GetLeft() + bbox.GetRight()) / 2
            center_y = (bbox.GetTop() + bbox.GetBottom()) / 2
            max_radius = min(bbox.GetWidth(), bbox.GetHeight()) / 2
            
            theta = 0
            radius = edge_clearance
            
            while radius <= max_radius:
                x = center_x + radius * math.cos(theta)
                y = center_y + radius * math.sin(theta)
                
                if randomize:
                    rand_theta = random.uniform(-0.1, 0.1)
                    rand_radius = random.uniform(-h_spacing * 0.1, h_spacing * 0.1)
                    x = center_x + (radius + rand_radius) * math.cos(theta + rand_theta)
                    y = center_y + (radius + rand_radius) * math.sin(theta + rand_theta)
                
                pos = pcbnew.VECTOR2I(int(x), int(y))
                
                if self.is_point_in_zone_polygon(pos, zone, zone_outline, edge_clearance):
                    positions.append(pos)
                
                theta += h_spacing / radius
                radius = edge_clearance + h_spacing * theta / (2 * math.pi)
        
        return positions
    
    def is_point_in_zone_polygon(self, point, zone, zone_outline, clearance):
        """ゾーンのポリゴン形状を使用して点がゾーン内にあるかチェック"""
        try:
            if zone_outline is not None:
                # SHAPE_POLY_SETを使用した正確な判定
                pt = pcbnew.VECTOR2I(int(point.x), int(point.y))
                if zone_outline.Contains(pt):
                    # クリアランスを考慮（境界からの距離をチェック）
                    # 簡易的に境界ボックスのマージンもチェック
                    bbox = zone.GetBoundingBox()
                    if (point.x >= bbox.GetLeft() + clearance and
                        point.x <= bbox.GetRight() - clearance and
                        point.y >= bbox.GetTop() + clearance and
                        point.y <= bbox.GetBottom() - clearance):
                        return True
                return False
            else:
                # フォールバック: BoundingBoxのみでチェック
                bbox = zone.GetBoundingBox()
                return (bbox.GetLeft() + clearance <= point.x <= bbox.GetRight() - clearance and 
                       bbox.GetTop() + clearance <= point.y <= bbox.GetBottom() - clearance)
        except Exception as e:
            print(f"ゾーン内チェックエラー: {e}")
            # エラー時はBoundingBoxでチェック
            try:
                bbox = zone.GetBoundingBox()
                return (bbox.GetLeft() + clearance <= point.x <= bbox.GetRight() - clearance and 
                       bbox.GetTop() + clearance <= point.y <= bbox.GetBottom() - clearance)
            except:
                return False
    
    def check_drc_fast(self, pos, via_size, via_clearance, spatial_index):
        """高速DRCチェック（空間インデックス使用、V2.2.0でvia_clearance対応）
        
        Args:
            pos: チェック位置
            via_size: VIAサイズ
            via_clearance: VIA間クリアランス
            spatial_index: 空間インデックス
        
        Returns:
            bool: 配置可能な場合True
        """
        # 最小距離 = VIAサイズ + クリアランス
        min_distance = via_size + via_clearance
        nearby_vias = spatial_index.get_nearby_vias(pos, min_distance)
        return len(nearby_vias) == 0
    
    def clear_vias(self, board, zones, net_code, settings):
        """VIAをクリア"""
        clear_plugin_vias = settings["clear_plugin_vias"]
        group_name = settings["group_name"]
        
        via_count = 0
        
        vias = []
        for item in board.GetTracks():
            if item.Type() == pcbnew.PCB_VIA_T:
                vias.append(item)
        
        progress_dlg = wx.ProgressDialog(
            "VIA削除中...",
            "VIAを削除しています...",
            maximum=len(vias),
            parent=None,
            style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE | wx.PD_CAN_ABORT | wx.PD_ELAPSED_TIME
        )
        
        try:
            group_vias = []
            if clear_plugin_vias:
                try:
                    for group in board.Groups():
                        try:
                            group_name_str = str(group.GetName())
                            if group_name_str.startswith(group_name):
                                for item in group.GetItems():
                                    if item.Type() == pcbnew.PCB_VIA_T:
                                        group_vias.append(item)
                        except Exception as e:
                            continue
                except:
                    pass
            
            processed = 0
            for zone in zones:
                bbox = zone.GetBoundingBox()
                
                for via in vias:
                    if processed % 10 == 0:
                        cont, skip = progress_dlg.Update(processed, f"VIA削除中... ({processed}/{len(vias)})")
                        if not cont:
                            wx.MessageBox(f"{via_count}個のVIAを削除して処理を中断しました", 
                                         "情報", wx.OK | wx.ICON_INFORMATION)
                            return
                    
                    try:
                        via_net_code = via.GetNetCode()
                        if via_net_code != net_code:
                            continue
                    except Exception as e:
                        continue
                    
                    if clear_plugin_vias and via not in group_vias:
                        continue
                    
                    via_pos = via.GetPosition()
                    if (bbox.GetLeft() <= via_pos.x <= bbox.GetRight() and
                        bbox.GetTop() <= via_pos.y <= bbox.GetBottom()):
                        board.Remove(via)
                        via_count += 1
                    
                    processed += 1
            
            wx.MessageBox(f"{via_count}個のVIAを削除しました", "完了", wx.OK | wx.ICON_INFORMATION)
        
        finally:
            progress_dlg.Destroy()
        
        pcbnew.Refresh()
    
    def add_via(self, board, pos, net_code, via_size, drill_size):
        """VIAを追加"""
        try:
            via = pcbnew.PCB_VIA(board)
            via.SetPosition(pos)
            via.SetWidth(via_size)
            via.SetDrill(drill_size)
            via.SetNetCode(net_code)
            via.SetLayerPair(pcbnew.F_Cu, pcbnew.B_Cu)
            board.Add(via)
            return via
        except Exception as e:
            print(f"VIA追加エラー: {e}")
            return None


# プラグインの登録
ViaStitchingPlugin().register()
