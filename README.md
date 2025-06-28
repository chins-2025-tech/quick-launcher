# QuickLauncherProject

Windows用クイックランチャー（Python/Tkinter）

## 概要

- Windows 11でWindows 10のタスクバーにある「ツールバー」の代替として開発しました。
- よく使うアプリ・ファイル・Webサイトをグループ分けしてトレイ（通知領域）のアイコンから素早く起動できるランチャーです。
- 設定・リンク情報はJSONファイルで管理。
- 編集画面・ポップアップUIはリサイズ追従・即時保存・システムアイコン統一。
- PyInstallerでスタンドアロン実行ファイル化も可能。

## 主な機能

- グループごとにリンク（アプリ/ファイル/URL）を管理
- 編集画面でドラッグ＆ドロップ並び替え、追加・削除・名前変更
- ポップアップUIでグループ→リンクを素早く選択
- システム・Webアイコン自動取得
- 設定画面でフォント・色・アイコン取得方法などカスタマイズ

## 使い方

1. Python 3.8以降が必要です。
2. 必要なパッケージをインストール：
   ```sh
   pip install -r requirements.txt
   ```
3. `quick_launcher.py` を実行：
   ```sh
   python quick_launcher.py
   ```
4. トレイ（通知領域）のアイコンをクリックして、よく使うアプリ・ファイル・Webサイトを素早く起動できます。
   また、トレイアイコンの右クリックメニューから「リンク編集」「設定」なども操作できます。

## ビルド（PyInstaller）

```sh
pyinstaller --noconsole --onefile --icon=icon.ico --add-data "icon.png;." quick_launcher.py
```

- 実行ファイル（exe）と同じフォルダに `icon.png` を配置すると、
  トレイアイコンやウィンドウアイコンとして自動的に `icon.png` が使用されます。
  （icoはPyInstaller用、pngはアプリ実行時の表示用です）

## 除外ファイル

- 仮想環境・ビルド生成物・settings.json/links.json などは `.gitignore` で除外しています。

## ライセンス

MIT License
