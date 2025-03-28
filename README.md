# 画像グリッド作成ツール

このアプリケーションは、複数の画像をグリッド形式で配置し、PDFとして出力することができるツールです。

## 機能

- 複数の画像をグリッド形式で配置
- A4/A3サイズのPDF出力
- グリッド線の表示/非表示
- グリッド線の色と太さのカスタマイズ
- 行と列のサイズ調整
- ドラッグ＆ドロップでの画像追加

## 必要条件

- Python 3.8以上
- macOS/Linux/Windows

## インストール方法

1. リポジトリをクローンまたはダウンロードします。

2. 起動スクリプトに実行権限を付与します：
   ```bash
   chmod +x start.sh
   ```

## 使用方法

1. 起動スクリプトを実行します：
   ```bash
   ./start.sh
   ```

2. アプリケーションが起動したら：
   - 「画像を追加」ボタンをクリックするか、画像ファイルをドラッグ＆ドロップして画像を追加
   - グリッド設定（行の高さ、列の幅、グリッド線の表示など）を調整
   - 「PDFを作成」ボタンをクリックしてPDFを出力

## 注意事項

- 対応している画像形式：PNG、JPG、JPEG
- 設定は自動的に保存され、次回起動時に復元されます
- 設定をリセットしたい場合は、メニューバーの「設定」→「設定をリセット」を選択してください 