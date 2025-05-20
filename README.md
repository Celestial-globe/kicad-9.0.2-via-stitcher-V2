# KiCAD VIA ステッチャープラグイン

このKiCADプラグインは、選択したゾーンに均等にVIAを配置し、オプションでグループ化する機能を提供します。

## 機能

- 選択したゾーンに均等にVIAを配置
- VIAのサイズ、ドリルサイズ、間隔をカスタマイズ可能
- 追加したVIAを自動的にグループ化するオプション
- ゾーンのネット名を使用するオプション
- 設定を保存して次回使用時のデフォルトとして使用可能

## インストール方法

1. このリポジトリをクローンまたはダウンロードします
2. ファイルを KiCAD のプラグインディレクトリに配置します:
   - Windows: `%APPDATA%\kicad\9.0\scripting\plugins\`
   - Linux: `~/.local/share/kicad/9.0/scripting/plugins/`
   - macOS: `~/Library/Application Support/kicad/9.0/scripting/plugins/`

## 使い方

1. KiCADのPCBエディタでVIAを配置したいゾーンを選択します
2. プラグインメニューから「VIAステッチ」を選択します
3. 必要な設定を行い、「OK」をクリックします

## 要件

- KiCAD 9.0.2 以上推奨