#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
KiCad VIAステッチプラグイン（最適化版 V2.0.7）
このプラグインは、KiCadのPCBエディタ(Pcbnew)で選択された銅エリアにVIAを自動配置します。
熱伝導や電流伝導を改善するためのVIAステッチングを効率的に行うことができます。
KiCad 9.0.2対応版 - 空間インデックス化とプログレスバー対応
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
PLUGIN_VERSION = "1.0.7"

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
    "group_name": "ViaStitching"  # VIAグループ名
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

# 設定ファイルのパス
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
                return json.load(f)
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
        wx.Dialog.__init__(self, parent, title="VIAステッチング v" + PLUGIN_VERSION, style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        
        self.board = board
        self.selected_area = selected_area
        self.settings = settings
        self.net_info = None
        self.net_data = []  # ネット名とIDのリスト
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
                self.net_info = 0  # デフォルトは未接続
        
        # UIの構築
        self.build_ui()
        
        # 初期値の設定
        self.set_initial_values()

        # アクションの設定
        self.action_radio.SetSelection(0)  # デフォルトは「配置」
        # 初期状態ではチェックボックスを無効化
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

    # ... [以下、既存のUIビルド関連メソッドは同じなので省略] ...
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
        param_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "パラメータ"), wx.VERTICAL)
        
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
        
        # クリアランス設定
        clearance_sizer = wx.BoxSizer(wx.HORIZONTAL)
        clearance_label = wx.StaticText(self, wx.ID_ANY, "エッジクリアランス (mm):")
        self.clearance_text = wx.TextCtrl(self, wx.ID_ANY)
        clearance_sizer.Add(clearance_label, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        clearance_sizer.Add(self.clearance_text, 1, wx.ALL|wx.EXPAND, 5)
        param_sizer.Add(clearance_sizer, 0, wx.EXPAND, 5)
        
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
        
        # プレビューエリア
        preview_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "プレビュー"), wx.VERTICAL)
        self.preview_panel = wx.Panel(self, wx.ID_ANY, size=(300, 200))
        self.preview_panel.SetBackgroundColour(wx.WHITE)
        self.preview_panel.Bind(wx.EVT_PAINT, self.on_paint_preview)
        preview_sizer.Add(self.preview_panel, 1, wx.ALL|wx.EXPAND, 5)
        main_sizer.Add(preview_sizer, 1, wx.ALL|wx.EXPAND, 5)
        
        # アクション選択用のサイザー
        action_sizer = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, "アクション"), wx.VERTICAL)

        # アクション選択
        action_choices = ["配置", "クリア"]
        self.action_radio = wx.RadioBox(self, wx.ID_ANY, "アクション", choices=action_choices, majorDimension=1)
        main_sizer.Add(self.action_radio, 0, wx.ALL|wx.EXPAND, 5)

        # プラグインVIAのみクリアオプション（アクションサイザー内に配置）
        self.clear_plugin_vias_checkbox = wx.CheckBox(self, wx.ID_ANY, "プラグインで配置したVIAのみをクリア")
        action_sizer.Add(self.clear_plugin_vias_checkbox, 0, wx.ALL, 5)
        main_sizer.Add(action_sizer, 0, wx.ALL|wx.EXPAND, 5)
        
        # ボタン
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.ok_button = wx.Button(self, wx.ID_OK, "OK")
        self.cancel_button = wx.Button(self, wx.ID_CANCEL, "キャンセル")
        self.clear_button = wx.Button(self, wx.ID_CLEAR, "クリア")
        button_sizer.Add(self.ok_button, 0, wx.ALL, 5)
        button_sizer.Add(self.cancel_button, 0, wx.ALL, 5)
        button_sizer.Add(self.clear_button, 0, wx.ALL, 5)
        main_sizer.Add(button_sizer, 0, wx.ALIGN_RIGHT, 5)
        
        self.SetSizer(main_sizer)
    
    def set_initial_values(self):
        """初期値の設定"""
        # ネットリストの設定
        self.populate_net_list()
        
        # パラメータの設定
        self.via_size_text.SetValue(str(self.settings["via_size"]))
        self.drill_size_text.SetValue(str(self.settings["drill_size"]))
        self.h_spacing_text.SetValue(str(self.settings["h_spacing"]))
        self.v_spacing_text.SetValue(str(self.settings["v_spacing"]))
        self.h_offset_text.SetValue(str(self.settings["h_offset"]))
        self.v_offset_text.SetValue(str(self.settings["v_offset"]))
        self.clearance_text.SetValue(str(self.settings["edge_clearance"]))
        
        # パターン選択
        pattern_index = 0
        for i, pattern in enumerate(self.pattern_data):
            if pattern["value"] == self.settings["pattern"]:
                pattern_index = i
                break
        self.pattern_choice.SetSelection(pattern_index)
        
        # チェックボックスの設定
        self.randomize_checkbox.SetValue(self.settings["randomize"])
        self.clear_plugin_vias_checkbox.SetValue(self.settings["clear_plugin_vias"])
        
        # アクションの設定
        self.action_radio.SetSelection(0)  # デフォルトは「配置」
    
    def populate_net_list(self):
        """ネットリストの設定"""
        self.net_choice.Clear()
        self.net_data = []
        
        try:
            # ボードからネットリストを取得
            nets = self.board.GetNetInfo().NetsByName()
            print(f"ネット数: {len(nets)}")
            
            # ネット名とIDのマッピングを作成
            for net_name, net_item in nets.items():
                # KiCadバージョン間のAPI互換性対応
                try:
                    # 新しいバージョン
                    net_id = net_item.GetNetCode()
                    print(f"ネット名: {net_name}, ID: {net_id}")
                except:
                    try:
                        # 中間バージョン
                        net_id = net_item.GetNet()
                    except:
                        try:
                            # 古いバージョン
                            net_id = net_item.GetCode()
                        except:
                            # どの方法も失敗した場合、インデックスを使用
                            net_id = 0
                
                # wxString型をPythonのstr型に変換
                self.net_data.append({"name": str(net_name), "id": net_id})
            
            # ネット名でソート
            self.net_data.sort(key=lambda x: x["name"])
            
            # Choiceコントロールに追加
            selected_index = 0
            for i, net in enumerate(self.net_data):
                self.net_choice.Append(net["name"])
                if net["id"] == self.net_info:
                    selected_index = i
            
            # 選択されたエリアのネットを選択
            if self.net_choice.GetCount() > 0:
                self.net_choice.SetSelection(selected_index)
                print(f"選択されたネット: {self.net_data[selected_index]['name']}")
        except Exception as e:
            print(f"ネットリスト取得エラー: {e}")
            traceback.print_exc()
            # エラーが発生した場合、未接続ネットを追加
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
        
        # プレビューエリアのサイズ
        width, height = self.preview_panel.GetSize()
        
        # 選択されたパターンに基づいてプレビューを描画
        pattern_index = self.pattern_choice.GetSelection()
        if pattern_index != wx.NOT_FOUND:
            pattern = self.pattern_data[pattern_index]["value"]
        else:
            pattern = "grid"  # デフォルト
        
        # パラメータの取得
        try:
            h_spacing = float(self.h_spacing_text.GetValue())
            v_spacing = float(self.v_spacing_text.GetValue())
            h_offset = float(self.h_offset_text.GetValue())
            v_offset = float(self.v_offset_text.GetValue())
            clearance = float(self.clearance_text.GetValue())
            randomize = self.randomize_checkbox.GetValue()
        except ValueError:
            return
        
        # スケーリング係数
        scale = min(width, height) / 20.0
        
        # 描画領域
        draw_width = width - 20
        draw_height = height - 20
        
        # 背景の描画
        dc.SetBrush(wx.Brush(wx.Colour(200, 200, 200)))
        dc.SetPen(wx.Pen(wx.Colour(100, 100, 100), 2))
        dc.DrawRectangle(10, 10, draw_width, draw_height)
        
        # VIAの描画
        dc.SetBrush(wx.Brush(wx.Colour(255, 255, 0)))
        dc.SetPen(wx.Pen(wx.Colour(0, 0, 0), 1))
        
        if pattern == "grid":
            # 格子状配置
            rows = int(draw_height / (v_spacing * scale))
            cols = int(draw_width / (h_spacing * scale))
            
            for row in range(rows):
                for col in range(cols):
                    x = 10 + col * h_spacing * scale + h_offset * scale
                    y = 10 + row * v_spacing * scale + v_offset * scale
                    
                    # エッジクリアランスのチェック
                    if (x >= 10 + clearance * scale and 
                        x <= 10 + draw_width - clearance * scale and 
                        y >= 10 + clearance * scale and 
                        y <= 10 + draw_height - clearance * scale):
                        
                        # ランダム化
                        if randomize:
                            x += random.uniform(-h_spacing * scale * 0.2, h_spacing * scale * 0.2)
                            y += random.uniform(-v_spacing * scale * 0.2, v_spacing * scale * 0.2)
                        
                        dc.DrawCircle(int(x), int(y), 3)
        
        elif pattern == "boundary":
            # 境界配置
            perimeter = 2 * (draw_width + draw_height)
            spacing = h_spacing * scale  # 境界に沿った間隔
            num_vias = int(perimeter / spacing)
            
            for i in range(num_vias):
                pos = i * spacing
                
                # 境界に沿って配置
                if pos < draw_width:
                    # 上辺
                    x = 10 + pos
                    y = 10 + clearance * scale
                elif pos < draw_width + draw_height:
                    # 右辺
                    x = 10 + draw_width - clearance * scale
                    y = 10 + (pos - draw_width)
                elif pos < 2 * draw_width + draw_height:
                    # 下辺
                    x = 10 + draw_width - (pos - draw_width - draw_height)
                    y = 10 + draw_height - clearance * scale
                else:
                    # 左辺
                    x = 10 + clearance * scale
                    y = 10 + draw_height - (pos - 2 * draw_width - draw_height)
                
                # ランダム化
                if randomize:
                    if pos < draw_width or (pos >= 2 * draw_width + draw_height):
                        # 上辺と左辺
                        x += random.uniform(-spacing * 0.1, spacing * 0.1)
                        y += random.uniform(0, spacing * 0.2)
                    else:
                        # 右辺と下辺
                        x -= random.uniform(0, spacing * 0.2)
                        y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                dc.DrawCircle(int(x), int(y), 3)
        
        elif pattern == "spiral":
            # スパイラル配置
            center_x = 10 + draw_width / 2
            center_y = 10 + draw_height / 2
            max_radius = min(draw_width, draw_height) / 2 - clearance * scale
            
            spacing = h_spacing * scale
            theta = 0
            radius = clearance * scale
            
            while radius <= max_radius:
                x = center_x + radius * math.cos(theta)
                y = center_y + radius * math.sin(theta)
                
                # ランダム化
                if randomize:
                    x += random.uniform(-spacing * 0.1, spacing * 0.1)
                    y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                dc.DrawCircle(int(x), int(y), 3)
                
                # スパイラルの次のポイント
                theta += spacing / radius
                radius = clearance * scale + spacing * theta / (2 * math.pi)
    
    def on_action_changed(self, event):
        """アクション変更時の処理"""
        # プレビューの更新
        # クリアアクションが選択された場合のみチェックボックスを表示
        if self.action_radio.GetSelection() == 1:  # "クリア"
            self.clear_plugin_vias_checkbox.Enable(True)
        else:
            self.clear_plugin_vias_checkbox.Enable(False)
        # レイアウトの更新
        self.Layout()
        self.update_preview()
    
    def on_randomize_changed(self, event):
        """ランダム化オプション変更時の処理"""
        # プレビューの更新
        self.update_preview()
    
    def on_pattern_changed(self, event):
        """パターン変更時の処理"""
        # プレビューの更新
        self.update_preview()
    
    def on_ok(self, event):
        """OKボタン押下時の処理"""
        try:
            # 設定の保存
            self.settings["via_size"] = float(self.via_size_text.GetValue())
            self.settings["drill_size"] = float(self.drill_size_text.GetValue())
            self.settings["h_spacing"] = float(self.h_spacing_text.GetValue())
            self.settings["v_spacing"] = float(self.v_spacing_text.GetValue())
            self.settings["h_offset"] = float(self.h_offset_text.GetValue())
            self.settings["v_offset"] = float(self.v_offset_text.GetValue())
            self.settings["edge_clearance"] = float(self.clearance_text.GetValue())
            
            # パターン
            pattern_index = self.pattern_choice.GetSelection()
            if pattern_index != wx.NOT_FOUND:
                self.settings["pattern"] = self.pattern_data[pattern_index]["value"]
            
            # チェックボックス
            self.settings["randomize"] = self.randomize_checkbox.GetValue()
            self.settings["clear_plugin_vias"] = self.clear_plugin_vias_checkbox.GetValue()
            
            # 設定の保存
            save_settings(self.settings)
            
            # ダイアログを閉じる
            self.EndModal(wx.ID_OK)
        except ValueError as e:
            wx.MessageBox(f"入力値が不正です: {e}", "エラー", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"エラーが発生しました: {e}", "エラー", wx.OK | wx.ICON_ERROR)
    
    def on_cancel(self, event):
        """キャンセルボタン押下時の処理"""
        # ダイアログを閉じる
        self.EndModal(wx.ID_CANCEL)
    
    def on_clear(self, event):
        """クリアボタン押下時の処理"""
        # デフォルト値に戻す
        self.via_size_text.SetValue(str(DEFAULT_SETTINGS["via_size"]))
        self.drill_size_text.SetValue(str(DEFAULT_SETTINGS["drill_size"]))
        self.h_spacing_text.SetValue(str(DEFAULT_SETTINGS["h_spacing"]))
        self.v_spacing_text.SetValue(str(DEFAULT_SETTINGS["v_spacing"]))
        self.h_offset_text.SetValue(str(DEFAULT_SETTINGS["h_offset"]))
        self.v_offset_text.SetValue(str(DEFAULT_SETTINGS["v_offset"]))
        self.clearance_text.SetValue(str(DEFAULT_SETTINGS["edge_clearance"]))
        
        # パターン選択
        for i, pattern in enumerate(self.pattern_data):
            if pattern["value"] == DEFAULT_SETTINGS["pattern"]:
                self.pattern_choice.SetSelection(i)
                break
        
        # チェックボックス
        self.randomize_checkbox.SetValue(DEFAULT_SETTINGS["randomize"])
        self.clear_plugin_vias_checkbox.SetValue(DEFAULT_SETTINGS["clear_plugin_vias"])
        
        # プレビューの更新
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
            # 無効な選択の場合は0（未接続）を返す
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
        self.description = "選択された銅エリアにVIAを自動配置します"
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), "via_stitcher_v2.png")
    
    def Run(self):
        """プラグインの実行"""
        try:
            # 現在のボードを取得
            board = pcbnew.GetBoard()
            print("ボード取得成功")
            
            # 選択されたアイテムを取得 - KiCadバージョン対応
            selected_zones = []
            
            # KiCad 9.0.2での選択ゾーン取得方法
            try:
                # 全てのゾーンを取得し、選択されているものをフィルタリング
                for zone in board.Zones():
                    if zone.IsSelected():
                        selected_zones.append(zone)
                        print(f"選択されたゾーン: {zone.GetNetname()}")
            except Exception as e:
                print(f"Zones()メソッドエラー: {e}")
                try:
                    # 代替方法1: 選択マネージャーを使用
                    selection = board.GetSelection()
                    for item in selection:
                        if item.IsType(pcbnew.PCB_ZONE_AREA_T):
                            selected_zones.append(item)
                            print("選択マネージャーからゾーンを取得")
                except Exception as e:
                    print(f"GetSelection()メソッドエラー: {e}")
                    try:
                        # 代替方法2: GetAreaCount/GetAreaを使用
                        for i in range(board.GetAreaCount()):
                            zone = board.GetArea(i)
                            if zone.IsSelected():
                                selected_zones.append(zone)
                                print(f"GetArea()からゾーンを取得: {i}")
                    except Exception as e:
                        print(f"GetArea()メソッドエラー: {e}")
            
            if not selected_zones:
                wx.MessageBox("銅エリア（ゾーン）を選択してください", "エラー", wx.OK | wx.ICON_ERROR)
                return
            
            print(f"選択されたゾーン数: {len(selected_zones)}")
            
            # 設定の読み込み
            settings = load_settings()
            
            # ダイアログの表示
            with ViaStitchingDialog(None, board, selected_zones[0], settings) as dlg:
                if dlg.ShowModal() == wx.ID_OK:
                    try:
                        action = dlg.get_action()
                        net_code = dlg.get_net_code()
                        print(f"アクション: {action}, ネットコード: {net_code}")
                        
                        # ネットコードの妥当性チェック
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
        """最適化されたVIA配置（空間インデックス + プログレスバー）"""
        # パラメータの取得
        via_size = pcbnew.FromMM(settings["via_size"])
        drill_size = pcbnew.FromMM(settings["drill_size"])
        h_spacing = pcbnew.FromMM(settings["h_spacing"])
        v_spacing = pcbnew.FromMM(settings["v_spacing"])
        h_offset = pcbnew.FromMM(settings["h_offset"])
        v_offset = pcbnew.FromMM(settings["v_offset"])
        edge_clearance = pcbnew.FromMM(settings["edge_clearance"])
        pattern = settings["pattern"]
        randomize = settings["randomize"]
        group_name = settings["group_name"]
        
        # 空間インデックス初期化（1.5倍のvia_sizeをグリッドサイズに使用）
        spatial_index = SpatialIndex(int(via_size * 1.5))
        
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
            positions = self.calculate_candidate_positions(zone, pattern, 
                                                         h_spacing, v_spacing, 
                                                         h_offset, v_offset, 
                                                         edge_clearance, randomize)
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
        
        try:
            # 有効な位置をフィルタリング
            valid_positions = []
            processed = 0
            
            for pos in all_candidate_positions:
                # プログレス更新
                if processed % 10 == 0:  # 10個ごとに更新
                    cont, skip = progress_dlg.Update(processed, f"位置検証中... ({processed}/{total_candidates})")
                    if not cont:  # キャンセルされた場合
                        wx.MessageBox("処理がキャンセルされました", "情報", wx.OK | wx.ICON_INFORMATION)
                        return
                
                # 高速DRCチェック
                if self.check_drc_fast(pos, via_size, spatial_index):
                    valid_positions.append(pos)
                
                processed += 1
            
            valid_count = len(valid_positions)
            print(f"有効位置数: {valid_count}")
            
            if valid_count == 0:
                wx.MessageBox("配置可能な位置がありません（DRCエラー）", "情報", wx.OK | wx.ICON_INFORMATION)
                return
            
            # プログレスダイアログを更新（VIA作成フェーズ）
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
                # プログレス更新
                if i % 5 == 0:  # 5個ごとに更新
                    cont, skip = progress_dlg.Update(i, f"VIA作成中... ({i}/{valid_count})")
                    if not cont:  # キャンセルされた場合
                        wx.MessageBox(f"{created_vias}個のVIAを配置して処理を中断しました", "情報", wx.OK | wx.ICON_INFORMATION)
                        return
                
                # VIAの作成
                via = self.add_via(board, pos, net_code, via_size, drill_size)
                if via:
                    # 空間インデックスに追加
                    spatial_index.add_via(pos, via)
                    
                    # グループに追加
                    if group:
                        try:
                            group.AddItem(via)
                        except Exception as e:
                            try:
                                group.AddPcbItem(via)
                            except:
                                pass
                    
                    created_vias += 1
            
            # 結果の表示
            wx.MessageBox(f"{created_vias}個のVIAを高速配置しました", "完了", wx.OK | wx.ICON_INFORMATION)
            
        finally:
            # プログレスダイアログを閉じる
            progress_dlg.Destroy()
        
        # ボードの更新
        pcbnew.Refresh()
    
    def calculate_candidate_positions(self, zone, pattern, h_spacing, v_spacing, 
                                    h_offset, v_offset, edge_clearance, randomize):
        """候補位置を事前計算"""
        positions = []
        bbox = zone.GetBoundingBox()
        
        if pattern == "grid":
            # 格子状配置
            x_start = bbox.GetLeft() + edge_clearance + h_offset
            y_start = bbox.GetTop() + edge_clearance + v_offset
            x_end = bbox.GetRight() - edge_clearance
            y_end = bbox.GetBottom() - edge_clearance
            
            x = x_start
            while x <= x_end:
                y = y_start
                while y <= y_end:
                    # ランダム化
                    if randomize:
                        rand_x = x + random.uniform(-h_spacing * 0.2, h_spacing * 0.2)
                        rand_y = y + random.uniform(-v_spacing * 0.2, v_spacing * 0.2)
                        pos = pcbnew.VECTOR2I(int(rand_x), int(rand_y))
                    else:
                        pos = pcbnew.VECTOR2I(int(x), int(y))
                    
                    # ゾーン内チェック
                    if self.is_point_in_zone_fast(pos, zone, edge_clearance):
                        positions.append(pos)
                    
                    y += v_spacing
                x += h_spacing
        
        elif pattern == "boundary":
            # 境界配置（簡易版）
            perimeter = 2 * (bbox.GetWidth() + bbox.GetHeight())
            spacing = h_spacing
            num_vias = int(perimeter / spacing)
            
            for i in range(num_vias):
                t = i / num_vias
                perimeter_pos = t * perimeter
                
                # 境界ボックスの周囲に沿って位置を計算
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
                
                # ランダム化
                if randomize:
                    rand_dist = random.uniform(0, spacing * 0.2)
                    if perimeter_pos < bbox.GetWidth() or (perimeter_pos >= 2 * bbox.GetWidth() + bbox.GetHeight()):
                        x += random.uniform(-spacing * 0.1, spacing * 0.1)
                        y += rand_dist
                    else:
                        x -= rand_dist
                        y += random.uniform(-spacing * 0.1, spacing * 0.1)
                
                pos = pcbnew.VECTOR2I(int(x), int(y))
                
                if self.is_point_in_zone_fast(pos, zone, edge_clearance):
                    positions.append(pos)
        
        elif pattern == "spiral":
            # スパイラル配置
            center_x = (bbox.GetLeft() + bbox.GetRight()) / 2
            center_y = (bbox.GetTop() + bbox.GetBottom()) / 2
            max_radius = min(bbox.GetWidth(), bbox.GetHeight()) / 2
            
            theta = 0
            radius = edge_clearance
            
            while radius <= max_radius:
                x = center_x + radius * math.cos(theta)
                y = center_y + radius * math.sin(theta)
                
                # ランダム化
                if randomize:
                    rand_theta = random.uniform(-0.1, 0.1)
                    rand_radius = random.uniform(-h_spacing * 0.1, h_spacing * 0.1)
                    x = center_x + (radius + rand_radius) * math.cos(theta + rand_theta)
                    y = center_y + (radius + rand_radius) * math.sin(theta + rand_theta)
                
                pos = pcbnew.VECTOR2I(int(x), int(y))
                
                if self.is_point_in_zone_fast(pos, zone, edge_clearance):
                    positions.append(pos)
                
                # スパイラルの次のポイント
                theta += h_spacing / radius
                radius = edge_clearance + h_spacing * theta / (2 * math.pi)
        
        return positions
    
    def check_drc_fast(self, pos, via_size, spatial_index):
        """高速DRCチェック（空間インデックス使用）"""
        min_distance = via_size * 1.1  # 少し余裕を持たせる
        
        # 空間インデックスを使用して近隣VIAを高速検索
        nearby_vias = spatial_index.get_nearby_vias(pos, min_distance)
        
        # 近くにVIAがある場合は配置不可
        return len(nearby_vias) == 0
    
    def is_point_in_zone_fast(self, point, zone, clearance):
        """高速なゾーン内判定"""
        try:
            bbox = zone.GetBoundingBox()
            return (bbox.GetLeft() + clearance <= point.x <= bbox.GetRight() - clearance and 
                   bbox.GetTop() + clearance <= point.y <= bbox.GetBottom() - clearance)
        except Exception as e:
            print(f"ゾーン内チェックエラー: {e}")
            return False
    
    def clear_vias(self, board, zones, net_code, settings):
        """VIAをクリア"""
        clear_plugin_vias = settings["clear_plugin_vias"]
        group_name = settings["group_name"]
        
        # 削除したVIAの数
        via_count = 0
        
        # ボード上のすべてのVIAを取得
        vias = []
        for item in board.GetTracks():
            if item.Type() == pcbnew.PCB_VIA_T:
                vias.append(item)
        
        # プログレスダイアログの作成
        progress_dlg = wx.ProgressDialog(
            "VIA削除中...",
            "VIAを削除しています...",
            maximum=len(vias),
            parent=None,
            style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE | wx.PD_CAN_ABORT | wx.PD_ELAPSED_TIME
        )
        
        try:
            # グループに属するVIAを取得
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
                            print(f"グループ名取得エラー: {e}")
                            continue
                except:
                    pass
            
            # 各ゾーンに対して処理
            processed = 0
            for zone in zones:
                bbox = zone.GetBoundingBox()
                
                for via in vias:
                    # プログレス更新
                    if processed % 10 == 0:
                        cont, skip = progress_dlg.Update(processed, f"VIA削除中... ({processed}/{len(vias)})")
                        if not cont:
                            wx.MessageBox(f"{via_count}個のVIAを削除して処理を中断しました", "情報", wx.OK | wx.ICON_INFORMATION)
                            return
                    
                    try:
                        via_net_code = via.GetNetCode()
                        if via_net_code != net_code:
                            continue
                    except Exception as e:
                        print(f"VIAネットコード取得エラー: {e}")
                        continue
                    
                    # プラグインで配置したVIAのみをクリアする場合
                    if clear_plugin_vias and via not in group_vias:
                        continue
                    
                    # VIAがゾーン内にあるか確認
                    if self.is_point_in_zone_fast(via.GetPosition(), zone, 0):
                        board.Remove(via)
                        via_count += 1
                    
                    processed += 1
            
            # 結果の表示
            wx.MessageBox(f"{via_count}個のVIAを削除しました", "完了", wx.OK | wx.ICON_INFORMATION)
        
        finally:
            progress_dlg.Destroy()
        
        # ボードの更新
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