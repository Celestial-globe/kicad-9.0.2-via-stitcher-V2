# -*- coding: utf-8 -*-

import pcbnew
import wx
import math
import os
import sys
import logging
import json  # JSONファイルの読み書きに使用

# ロギング設定
logging.basicConfig(filename=os.path.join(os.path.dirname(__file__), 'via_stitcher.log'),
                    level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

class ViaStitcher(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "VIAステッチ"
        self.category = "配置"
        self.description = "選択したゾーンにVIAを均等に配置し、オプションでグループ化します"
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(os.path.dirname(__file__), 'via_stitch_icon.png')
        
    def Run(self):
        try:
            # 処理開始時刻を記録（パフォーマンス計測用）
            import time
            start_time = time.time()
            
            board = pcbnew.GetBoard()
            logging.debug("Starting ViaStitcher plugin")
            
            # 選択されているゾーンを取得
            selected_zones = []
            for zone in board.Zones():
                if zone.IsSelected():
                    selected_zones.append(zone)
                    logging.debug(f"Selected zone: {zone.GetNetname()}")
            
            # 選択されたゾーンがない場合
            if not selected_zones:
                wx.MessageBox("ゾーンが選択されていません。", "情報", wx.OK | wx.ICON_INFORMATION)
                return
                
            # 最初に選択されたゾーンのネット名を取得
            selected_zone_net = selected_zones[0].GetNetname() if selected_zones else ""
            
            # 進捗ダイアログ用の変数
            progress_dialog = None
            
            # ダイアログを表示して設定を取得（選択されたゾーン名を渡す）
            dialog = ViaStitcherDialog(None, selected_zone_net)
            if dialog.ShowModal() == wx.ID_OK:
                # ユーザー設定を取得
                via_size = pcbnew.FromMM(dialog.via_size_mm)
                via_drill = pcbnew.FromMM(dialog.via_drill_mm)
                via_spacing = pcbnew.FromMM(dialog.via_spacing_mm)
                net_name = dialog.net_name
                group_vias = dialog.group_vias  # グループ化設定を取得
                use_zone_net = dialog.use_zone_net  # ゾーンのネット名を使用するかどうか
                
                # 設定を保存
                if dialog.save_settings:
                    self.save_settings(dialog.via_size_mm, dialog.via_drill_mm, dialog.via_spacing_mm, 
                                      net_name, group_vias, use_zone_net)
                
                logging.debug(f"Settings: via_size={dialog.via_size_mm}mm, via_drill={dialog.via_drill_mm}mm, via_spacing={dialog.via_spacing_mm}mm, net_name={net_name}, group_vias={group_vias}, use_zone_net={use_zone_net}")
                
                # 進捗ダイアログを作成
                progress_dialog = wx.ProgressDialog("VIAステッチ", 
                                               "VIAを配置中...", 
                                               maximum=len(selected_zones),
                                               parent=None,
                                               style=wx.PD_AUTO_HIDE | wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME)
                
                # 選択されているアイテムを取得（ゾーンのみ使用）
                selected_items = selected_zones
                
                # 追加したVIAを記録するためのリスト
                added_vias = []
                # ネットごとのVIAを分けて保存するための辞書
                net_vias_map = {}
                # 選択されたアイテムにVIAを配置
                vias_count = 0
                
                for idx, item in enumerate(selected_items):
                    if isinstance(item, pcbnew.ZONE):
                        # 進捗状況を更新
                        if progress_dialog:
                            progress_dialog.Update(idx, f"ゾーン {idx+1}/{len(selected_items)} にVIAを配置中...")
                            
                        # ゾーンのネット名を使用する場合、その都度ネットを取得
                        current_net_name = item.GetNetname() if use_zone_net else net_name
                        logging.debug(f"Using net for zone: {current_net_name}")
                        
                        # ネット名からネットコードを取得
                        net_code = self.get_net_code(board, current_net_name)
                        
                        if net_code is not None:
                            zone_vias = self.place_vias_in_zone(board, item, via_size, via_drill, via_spacing, net_code)
                            # 追加したVIAを全体のリストに追加
                            added_vias.extend(zone_vias)
                            # ネットごとのマップにも追加
                            if current_net_name not in net_vias_map:
                                net_vias_map[current_net_name] = []
                            net_vias_map[current_net_name].extend(zone_vias)
                            
                            vias_count += len(zone_vias)
                        else:
                            wx.MessageBox(f"ネット '{current_net_name}' が見つかりません。", "エラー", wx.OK | wx.ICON_ERROR)
                
                # 進捗ダイアログを更新（VIAのグループ化処理）
                if progress_dialog and group_vias and added_vias:
                    progress_dialog.Update(len(selected_items), "VIAをグループ化中...")
                
                if vias_count > 0:
                    # グループ化オプションが有効な場合、追加したVIAをグループ化
                    # ネットごとに異なるグループにする
                    if group_vias and added_vias:
                        if use_zone_net:
                            # ネットごとに別々のグループを作成
                            for net_name, vias in net_vias_map.items():
                                if vias:  # 空のリストでなければグループ化
                                    group_name = f"VIA Stitch - {net_name}"
                                    self.create_via_group(board, vias, group_name)
                                    logging.debug(f"Created group for net {net_name} with {len(vias)} vias")
                        else:
                            # 単一のネットの場合は一つのグループにまとめる
                            group_name = f"VIA Stitch - {net_name}"
                            self.create_via_group(board, added_vias, group_name)
                            logging.debug(f"Created single group with {len(added_vias)} vias")
                        
                    # 処理時間を計算
                    elapsed_time = time.time() - start_time
                    logging.debug(f"Total processing time: {elapsed_time:.2f} seconds")
                    
                    # 変更を反映
                    pcbnew.Refresh()
                    message = f"VIAの配置が完了しました。{vias_count}個のVIAを追加しました。"
                    if group_vias:
                        if use_zone_net and len(net_vias_map) > 1:
                            message += f"\n{len(net_vias_map)}個のグループに分けてVIAがグループ化されました。"
                        else:
                            message += "\nVIAはグループ化されました。"
                    message += f"\n処理時間: {elapsed_time:.2f}秒"
                    wx.MessageBox(message, "完了", wx.OK | wx.ICON_INFORMATION)
                else:
                    wx.MessageBox("VIAを配置できませんでした。", "情報", wx.OK | wx.ICON_INFORMATION)
            
            dialog.Destroy()
            if progress_dialog:
                progress_dialog.Destroy()
            
        except Exception as e:
            logging.error(f"Error in Run method: {e}")
            import traceback
            logging.error(traceback.format_exc())
            if 'progress_dialog' in locals() and progress_dialog:
                progress_dialog.Destroy()
            wx.MessageBox(f"エラーが発生しました: {str(e)}", "エラー", wx.OK | wx.ICON_ERROR).Destroy()
            
        except Exception as e:
            logging.error(f"Error in Run method: {e}")
            import traceback
            logging.error(traceback.format_exc())
            wx.MessageBox(f"エラーが発生しました: {str(e)}", "エラー", wx.OK | wx.ICON_ERROR)

    def get_net_code(self, board, net_name):
        """ネット名からネットコードを取得するメソッド"""
        net_code = None
        try:
            # KiCAD 9.0.2での新しいアプローチを試みる
            netinfo = board.GetNetInfo()
            logging.debug(f"Attempting to find net '{net_name}'")
            
            # アプローチ1: FindNet (一部のバージョンでは使用できない)
            if hasattr(netinfo, 'FindNet'):
                net = netinfo.FindNet(net_name)
                if net:
                    net_code = net.GetNetCode()
                    logging.debug(f"Found net using FindNet: {net_name}, code: {net_code}")
            # アプローチ2: GetNetItem
            elif hasattr(netinfo, 'GetNetItem'):
                net = netinfo.GetNetItem(net_name)
                if net:
                    net_code = net.GetNetCode()
                    logging.debug(f"Found net using GetNetItem: {net_name}, code: {net_code}")
            # アプローチ3: 全ネット列挙
            else:
                for net_item in netinfo.NetsByNetcode():
                    net_obj = net_item[1]  # インデックス1にネットオブジェクトがある
                    netname = net_obj.GetNetname()
                    logging.debug(f"Found net: {netname}")
                    if netname == net_name:
                        net_code = net_item[0]  # インデックス0にネットコードがある
                        logging.debug(f"Found net in list: {net_name}, code: {net_code}")
                        break
                        
            # アプローチ4: GetNets (古いバージョン)
            if net_code is None and hasattr(board, 'GetNetsByName'):
                nets_by_name = board.GetNetsByName()
                if net_name in nets_by_name:
                    net_obj = nets_by_name[net_name]
                    net_code = net_obj.GetNetCode()
                    logging.debug(f"Found net using GetNetsByName: {net_name}, code: {net_code}")
        
        except Exception as e:
            logging.error(f"Error finding net: {e}")
            import traceback
            logging.error(traceback.format_exc())
        
        return net_code

    def place_vias_in_zone(self, board, zone, via_size, via_drill, via_spacing, net_code):
        added_vias = []  # 追加されたVIAを記録するリスト
        try:
            # パフォーマンス向上のため、既存のVIAの位置をあらかじめ収集
            existing_via_positions = self.collect_existing_via_positions(board)
            logging.debug(f"Collected {len(existing_via_positions)} existing via positions")
            
            # ゾーンの境界ボックスを取得
            bbox = zone.GetBoundingBox()
            
            # 境界ボックスの範囲内でVIAを配置
            x_start = bbox.GetX()
            x_end = bbox.GetX() + bbox.GetWidth()
            y_start = bbox.GetY()
            y_end = bbox.GetY() + bbox.GetHeight()
            
            # VIAを配置する間隔を計算
            x_count = max(1, int((x_end - x_start) / via_spacing))
            y_count = max(1, int((y_end - y_start) / via_spacing))
            
            # 実際の間隔を再計算
            x_step = (x_end - x_start) / x_count
            y_step = (y_end - y_start) / y_count
            
            # デバッグ情報をログに記録
            logging.debug(f"Zone bounding box: ({x_start}, {y_start}) to ({x_end}, {y_end})")
            logging.debug(f"Via spacing: {via_spacing}, X count: {x_count}, Y count: {y_count}")
            
            # ゾーンのアウトラインを取得（可能であれば一度だけ）
            zone_checker = self.get_zone_checker(zone)
            
            # ネット情報をキャッシュ
            netinfo = board.GetNetInfo()
            net = None
            if hasattr(netinfo, 'GetNetItem'):
                net = netinfo.GetNetItem(net_code)
                if net:
                    logging.debug(f"Using net code {net_code} for net '{net.GetNetname()}'")
            
            # 各位置にVIAを一括して配置（下準備）
            via_positions = []
            for i in range(x_count + 1):
                for j in range(y_count + 1):
                    x = int(x_start + i * x_step)
                    y = int(y_start + j * y_step)
                    
                    # 指定した位置がゾーン内かチェック
                    point = pcbnew.VECTOR2I(x, y)
                    
                    # 既存のVIAがないかチェック（高速化）
                    position_key = (x, y)
                    if position_key in existing_via_positions:
                        continue
                    
                    # ゾーン内部かチェック
                    if zone_checker(point):
                        via_positions.append(point)
            
            # 実際にVIAを作成（一括処理）
            logging.debug(f"Creating {len(via_positions)} vias")
            for point in via_positions:
                try:
                    # VIAを作成
                    via = pcbnew.PCB_VIA(board)
                    via.SetPosition(point)
                    via.SetWidth(via_size)
                    via.SetDrill(via_drill)
                    
                    # ネットコードを設定（最適化）
                    try:
                        via.SetNetCode(net_code)
                    except Exception as e:
                        logging.debug(f"SetNetCode failed, trying alternative: {e}")
                        # 代替方法
                        if hasattr(via, 'SetNet') and net:
                            via.SetNet(net)
                    
                    via.SetViaType(pcbnew.VIATYPE_THROUGH)
                    
                    # 基板に追加
                    board.Add(via)
                    # 追加したVIAをリストに記録
                    added_vias.append(via)
                except Exception as e:
                    logging.error(f"Error adding via at ({point.x}, {point.y}): {e}")
            
        except Exception as e:
            logging.error(f"Error in placing vias: {e}")
            import traceback
            logging.error(traceback.format_exc())
            
        return added_vias  # 追加したVIAのリストを返す
        
    # 既存のVIA位置をすべて収集する（高速化のため）
    def collect_existing_via_positions(self, board, tolerance=100):
        positions = set()
        for item in board.GetTracks():
            if type(item) == pcbnew.PCB_VIA:
                via_pos = item.GetPosition()
                x = via_pos.x
                y = via_pos.y
                # 許容誤差を含めたグリッドにスナップする
                x_grid = int(x / tolerance) * tolerance
                y_grid = int(y / tolerance) * tolerance
                # 周囲のグリッドセルも追加（許容誤差を考慮）
                for dx in range(-1, 2):
                    for dy in range(-1, 2):
                        positions.add((x_grid + dx * tolerance, y_grid + dy * tolerance))
        return positions
        
    # ゾーン内部判定のためのチェッカー関数を作成
    def get_zone_checker(self, zone):
        """ゾーン内部判定のための最適化された関数を返す"""
        try:
            # アプローチ1: GetFilledPolys を使用
            if hasattr(zone, 'GetFilledPolys'):
                filled_polys = zone.GetFilledPolys()
                if filled_polys:
                    def checker(point):
                        for poly in filled_polys:
                            if poly.Contains(point):
                                return True
                        return False
                    return checker
                    
            # アプローチ2: Outline を使用
            if hasattr(zone, 'GetOutline'):
                outline = zone.GetOutline()
                def checker(point):
                    return outline.Contains(point)
                return checker
                
            # アプローチ3: GetOutlines を使用
            if hasattr(zone, 'GetOutlines'):
                outlines = zone.GetOutlines()
                if outlines:
                    def checker(point):
                        for outline in outlines:
                            if outline.Contains(point):
                                return True
                        return False
                    return checker
                    
            # アプローチ4: Polygon を使用
            if hasattr(zone, 'GetPolygon'):
                polygon = zone.GetPolygon()
                def checker(point):
                    return polygon.Contains(point)
                return checker
                
            # どれも使えない場合は常にTrueを返す単純な関数
            logging.debug("Using simple bounding box check for zone")
            return lambda point: True
            
        except Exception as e:
            logging.error(f"Error creating zone checker: {e}")
            # エラーが発生した場合は単純な境界ボックスチェック
            return lambda point: True

    # 追加されたVIAをグループ化するメソッド
    def create_via_group(self, board, vias, group_name_base):
        try:
            # 一意のグループ名を生成（タイムスタンプを追加）
            import time
            timestamp = time.strftime("%Y%m%d%H%M%S")
            group_name = f"{group_name_base}_{timestamp}"
            
            # 既存のグループ名をチェックして、必要なら連番を追加
            existing_groups = self.get_existing_group_names(board)
            if group_name in existing_groups:
                # 同名のグループが既に存在する場合は連番を付ける
                counter = 1
                while f"{group_name}_{counter}" in existing_groups:
                    counter += 1
                group_name = f"{group_name}_{counter}"
            
            logging.debug(f"Using unique group name: {group_name}")
            
            # KiCAD 9.0.2 でのグループ作成方法
            # PCB_GROUP クラスを使用
            try:
                # まず、このバージョンでグループ化がサポートされているか確認
                if hasattr(pcbnew, 'PCB_GROUP'):  # PCB_GROUP クラスが存在するか確認
                    group = pcbnew.PCB_GROUP(board)
                    group.SetName(group_name)
                    
                    # グループにVIAを追加
                    for via in vias:
                        group.AddItem(via)
                    
                    # ボードにグループを追加
                    board.Add(group)
                    logging.debug(f"Created group '{group_name}' with {len(vias)} vias")
                    return True
                else:
                    # 代替方法: PCBNEW_SELECTION クラスを試す
                    if hasattr(pcbnew, 'PCBNEW_SELECTION'):
                        selection = pcbnew.PCBNEW_SELECTION()
                        for via in vias:
                            selection.Add(via)
                            
                        # グループ名を設定
                        selection.SetName(group_name)
                        
                        # ボードに追加
                        board.Add(selection)
                        logging.debug(f"Created selection group '{group_name}' with {len(vias)} vias")
                        return True
                    
                    logging.warning("Grouping not supported in this version of KiCAD")
                    return False
            except Exception as e:
                logging.error(f"Error creating group: {e}")
                import traceback
                logging.error(traceback.format_exc())
                return False
        except Exception as e:
            logging.error(f"Error in group creation: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False
            
    # 既存のグループ名をすべて取得するメソッド
    def get_existing_group_names(self, board):
        existing_names = set()
        try:
            # KiCAD 9.0.2 でのグループ取得方法
            if hasattr(board, 'Groups'):
                for group in board.Groups():
                    existing_names.add(group.GetName())
            # 古いバージョン用
            elif hasattr(pcbnew, 'PCB_GROUP') and hasattr(board, 'GetAllGroups'):
                for group in board.GetAllGroups():
                    existing_names.add(group.GetName())
                    
            logging.debug(f"Found {len(existing_names)} existing group names")
            return existing_names
        except Exception as e:
            logging.error(f"Error getting existing group names: {e}")
            return existing_names
    
    # 設定を保存するメソッド
    def save_settings(self, via_size_mm, via_drill_mm, via_spacing_mm, net_name, group_vias, use_zone_net=False):
        try:
            # 設定を辞書形式で準備
            settings = {
                'via_size_mm': via_size_mm,
                'via_drill_mm': via_drill_mm,
                'via_spacing_mm': via_spacing_mm,
                'net_name': net_name,
                'group_vias': group_vias,
                'use_zone_net': use_zone_net  # 新しい設定を追加
            }
            
            # 設定ファイルのパスを取得
            settings_path = os.path.join(os.path.dirname(__file__), 'via_stitcher_settings.json')
            
            # JSON形式で設定を保存
            with open(settings_path, 'w') as f:
                json.dump(settings, f)
            
            logging.debug(f"Settings saved to {settings_path}")
            return True
        except Exception as e:
            logging.error(f"Error saving settings: {e}")
            return False
    
    # 設定を読み込むメソッド
    def load_settings(self):
        try:
            # 設定ファイルのパスを取得
            settings_path = os.path.join(os.path.dirname(__file__), 'via_stitcher_settings.json')
            
            # ファイルが存在する場合のみ読み込み
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    settings = json.load(f)
                
                logging.debug(f"Settings loaded from {settings_path}")
                return settings
            else:
                logging.debug("Settings file not found, using defaults")
                return None
        except Exception as e:
            logging.error(f"Error loading settings: {e}")
            return None


class ViaStitcherDialog(wx.Dialog):
    def __init__(self, parent, selected_zone_net=""):
        wx.Dialog.__init__(self, parent, id=wx.ID_ANY, title="VIAステッチ設定", size=(300, 380))
        
        # デフォルト値
        self.via_size_mm = 0.6
        self.via_drill_mm = 0.3
        self.via_spacing_mm = 2.0
        self.net_name = "GND"
        self.group_vias = True
        self.save_settings = False  # 初期値は設定を保存しない
        self.use_zone_net = False   # 新しいオプション: ゾーンのネット名を使用するかどうか
        self.selected_zone_net = selected_zone_net  # 選択されたゾーンのネット名
        
        # 保存された設定を読み込む
        plugin = ViaStitcher()
        settings = plugin.load_settings()
        if settings:
            self.via_size_mm = settings.get('via_size_mm', self.via_size_mm)
            self.via_drill_mm = settings.get('via_drill_mm', self.via_drill_mm)
            self.via_spacing_mm = settings.get('via_spacing_mm', self.via_spacing_mm)
            self.net_name = settings.get('net_name', self.net_name)
            self.group_vias = settings.get('group_vias', self.group_vias)
            self.use_zone_net = settings.get('use_zone_net', self.use_zone_net)  # 新しい設定を読み込む
        
        # レイアウト作成
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 選択されたゾーンのネット名を表示（新機能）
        if selected_zone_net:
            zone_net_sizer = wx.BoxSizer(wx.HORIZONTAL)
            zone_net_label = wx.StaticText(self, label="選択ゾーンのネット名:")
            zone_net_value = wx.StaticText(self, label=selected_zone_net)
            zone_net_sizer.Add(zone_net_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
            zone_net_sizer.Add(zone_net_value, 1, wx.ALL | wx.EXPAND, 5)
            main_sizer.Add(zone_net_sizer, 0, wx.EXPAND)
            
            # 線で区切る
            main_sizer.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.ALL, 5)
        
        # VIAサイズ
        size_sizer = wx.BoxSizer(wx.HORIZONTAL)
        size_label = wx.StaticText(self, label="VIAサイズ (mm):")
        self.size_ctrl = wx.TextCtrl(self, value=str(self.via_size_mm))
        size_sizer.Add(size_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        size_sizer.Add(self.size_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(size_sizer, 0, wx.EXPAND)
        
        # VIAドリル
        drill_sizer = wx.BoxSizer(wx.HORIZONTAL)
        drill_label = wx.StaticText(self, label="VIAドリル (mm):")
        self.drill_ctrl = wx.TextCtrl(self, value=str(self.via_drill_mm))
        drill_sizer.Add(drill_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        drill_sizer.Add(self.drill_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(drill_sizer, 0, wx.EXPAND)
        
        # VIA間隔
        spacing_sizer = wx.BoxSizer(wx.HORIZONTAL)
        spacing_label = wx.StaticText(self, label="VIA間隔 (mm):")
        self.spacing_ctrl = wx.TextCtrl(self, value=str(self.via_spacing_mm))
        spacing_sizer.Add(spacing_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        spacing_sizer.Add(self.spacing_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(spacing_sizer, 0, wx.EXPAND)
        
        # ネット名
        net_sizer = wx.BoxSizer(wx.HORIZONTAL)
        net_label = wx.StaticText(self, label="ネット名:")
        self.net_ctrl = wx.TextCtrl(self, value=self.net_name)
        net_sizer.Add(net_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        net_sizer.Add(self.net_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(net_sizer, 0, wx.EXPAND)
        
        # ゾーンのネット名を使用するオプション（新機能）
        zone_net_option_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.zone_net_checkbox = wx.CheckBox(self, label="ゾーンのネット名を使用する")
        self.zone_net_checkbox.SetValue(self.use_zone_net)
        zone_net_option_sizer.Add(self.zone_net_checkbox, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(zone_net_option_sizer, 0, wx.EXPAND)
        
        # ゾーンネット名の使用オプションが変更されたときのイベント
        self.zone_net_checkbox.Bind(wx.EVT_CHECKBOX, self.on_zone_net_option_changed)
        
        # VIAをグループ化するオプション
        group_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.group_checkbox = wx.CheckBox(self, label="追加したVIAをグループ化する")
        self.group_checkbox.SetValue(self.group_vias)
        group_sizer.Add(self.group_checkbox, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(group_sizer, 0, wx.EXPAND)
        
        # 設定を保存するオプション
        save_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.save_checkbox = wx.CheckBox(self, label="この設定を次回のデフォルトとして保存する")
        self.save_checkbox.SetValue(self.save_settings)
        save_sizer.Add(self.save_checkbox, 1, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(save_sizer, 0, wx.EXPAND)
        
        # OKとキャンセルボタン
        button_sizer = wx.StdDialogButtonSizer()
        ok_button = wx.Button(self, wx.ID_OK)
        ok_button.SetDefault()
        button_sizer.AddButton(ok_button)
        cancel_button = wx.Button(self, wx.ID_CANCEL)
        button_sizer.AddButton(cancel_button)
        button_sizer.Realize()
        main_sizer.Add(button_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        
        self.SetSizer(main_sizer)
        
        # イベントバインド
        self.Bind(wx.EVT_BUTTON, self.on_ok, ok_button)
        
        # 初期状態でUI要素の有効/無効を設定
        self.update_ui_state()
        
    def on_zone_net_option_changed(self, event):
        """ゾーンネット名の使用オプションが変更されたときのハンドラ"""
        self.update_ui_state()
        
    def update_ui_state(self):
        """UI要素の有効/無効状態を更新"""
        use_zone_net = self.zone_net_checkbox.GetValue()
        # ゾーンネットを使用する場合はネット名入力欄を無効に
        self.net_ctrl.Enable(not use_zone_net)
        
    def on_ok(self, event):
        try:
            self.via_size_mm = float(self.size_ctrl.GetValue())
            self.via_drill_mm = float(self.drill_ctrl.GetValue())
            self.via_spacing_mm = float(self.spacing_ctrl.GetValue())
            self.net_name = self.net_ctrl.GetValue()
            self.group_vias = self.group_checkbox.GetValue()
            self.save_settings = self.save_checkbox.GetValue()  # 設定を保存するかどうか
            self.use_zone_net = self.zone_net_checkbox.GetValue()  # ゾーンのネット名を使用するかどうか
            
            if self.via_size_mm <= 0 or self.via_drill_mm <= 0 or self.via_spacing_mm <= 0:
                wx.MessageBox("値は正の数を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)
                return
                
            if self.via_drill_mm >= self.via_size_mm:
                wx.MessageBox("ドリルサイズはVIAサイズより小さくしてください。", "エラー", wx.OK | wx.ICON_ERROR)
                return
                
            event.Skip()  # ダイアログを閉じる
        
        except ValueError:
            wx.MessageBox("有効な数値を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)


# プラグインの登録
ViaStitcher().register()