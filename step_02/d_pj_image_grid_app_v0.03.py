import json
import logging
import os
import sys
import tempfile
from dataclasses import asdict, dataclass, field
from functools import lru_cache
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image, UnidentifiedImageError
from PyQt6.QtCore import QRectF, QSize, Qt, QThread, pyqtSignal
from PyQt6.QtGui import (QColor, QDragEnterEvent, QDropEvent, QImage, QPainter,
                         QPen, QPixmap)
from PyQt6.QtWidgets import (QApplication, QCheckBox, QColorDialog, QComboBox,
                             QDoubleSpinBox, QFileDialog, QFrame, QGridLayout,
                             QLabel, QMessageBox, QProgressDialog, QPushButton,
                             QScrollArea, QSpinBox, QVBoxLayout, QWidget)
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas

# ロギングの設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 定数定義
MM_TO_PT: float = 2.83465  # ミリメートルからポイントへの変換係数
DEFAULT_ROW_HEIGHT_MM: float = 100.0
DEFAULT_COL_WIDTH_MM: float = 100.0
DEFAULT_GRID_WIDTH: int = 1
DEFAULT_PREVIEW_HEIGHT: int = 600
THUMBNAIL_SIZE: Tuple[int, int] = (200, 200)  # プレビュー用のサムネイルサイズ
SETTINGS_FILE: str = "grid_settings.json"  # 設定ファイルのパス


@dataclass
class GridSettings:
    """グリッド設定を保持するデータクラス"""
    row_height_mm: float = DEFAULT_ROW_HEIGHT_MM
    col_width_mm: float = DEFAULT_COL_WIDTH_MM
    grid_line_visible: bool = True
    grid_color: QColor = field(default_factory=lambda: QColor(0, 0, 0))
    grid_width: int = DEFAULT_GRID_WIDTH
    page_size: Tuple[float, float] = A4

    def to_dict(self) -> Dict[str, Any]:
        """設定を辞書形式に変換"""
        settings_dict = asdict(self)
        # QColorをRGB値のタプルに変換
        settings_dict['grid_color'] = self.grid_color.getRgb()
        # ページサイズを文字列に変換
        settings_dict['page_size'] = 'A4' if self.page_size == A4 else 'A3'
        return settings_dict

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'GridSettings':
        """辞書から設定を復元"""
        # QColorを復元
        if 'grid_color' in data:
            color_data = data['grid_color']
            data['grid_color'] = QColor(*color_data)
        
        # ページサイズを復元
        if 'page_size' in data:
            data['page_size'] = A4 if data['page_size'] == 'A4' else A3
        
        return cls(**data)

    def save_to_file(self, file_path: str = SETTINGS_FILE) -> None:
        """設定をファイルに保存"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, indent=4)
            logger.info(f"設定を保存しました: {file_path}")
        except Exception as e:
            logger.error(f"設定の保存中にエラーが発生しました: {e}")
            raise

    @classmethod
    def load_from_file(cls, file_path: str = SETTINGS_FILE) -> 'GridSettings':
        """設定をファイルから読み込み"""
        try:
            if not os.path.exists(file_path):
                logger.info(f"設定ファイルが存在しません: {file_path}")
                return cls()
            
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except Exception as e:
            logger.error(f"設定の読み込み中にエラーが発生しました: {e}")
            return cls()


class PDFGenerationThread(QThread):
    """PDF生成をバックグラウンドで実行するスレッド"""
    finished = pyqtSignal(str, str)  # 成功時に一時ファイルパスとディレクトリを送信
    error = pyqtSignal(str)     # エラー時にメッセージを送信
    progress = pyqtSignal(int)  # 進捗状況を送信

    def __init__(self, image_paths: List[str], settings: GridSettings):
        super().__init__()
        self.image_paths = image_paths
        self.settings = settings
        self.temp_dir = None

    def run(self) -> None:
        try:
            self.temp_dir = tempfile.mkdtemp()  # 一時ディレクトリを作成
            file_path = os.path.join(self.temp_dir, "output.pdf")
            pdf = canvas.Canvas(file_path, pagesize=self.settings.page_size)
            
            page_width, page_height = self.settings.page_size
            
            # 行と列の数を計算
            col_width_pt = self.settings.col_width_mm * MM_TO_PT
            row_height_pt = self.settings.row_height_mm * MM_TO_PT
            cols = max(1, int(page_width / col_width_pt))
            rows = max(1, int(page_height / row_height_pt))
            
            total_cells = rows * cols
            processed_cells = 0

            # 画像の配置
            for row in range(rows):
                for col in range(cols):
                    cell_index = row * cols + col
                    img_index = cell_index % len(self.image_paths) if self.image_paths else 0
                    
                    if self.image_paths:
                        try:
                            self._process_image(pdf, self.image_paths[img_index], 
                                              row, col, col_width_pt, row_height_pt,
                                              page_height, self.temp_dir)
                        except UnidentifiedImageError as e:
                            logger.error(f"画像の読み込みに失敗しました: {self.image_paths[img_index]}, エラー: {e}")
                            continue
                        except Exception as e:
                            logger.error(f"画像の処理中にエラーが発生しました: {self.image_paths[img_index]}, エラー: {e}")
                            continue
                    
                    processed_cells += 1
                    progress = int((processed_cells / total_cells) * 100)
                    self.progress.emit(progress)
            
            # グリッド線の描画
            if self.settings.grid_line_visible:
                self._draw_grid_lines(pdf, cols, rows, col_width_pt, row_height_pt,
                                    page_width, page_height)
            
            pdf.save()
            self.finished.emit(file_path, self.temp_dir)
            
        except Exception as e:
            logger.error(f"PDF生成中にエラーが発生しました: {e}")
            self.error.emit(str(e))

    def __del__(self):
        """デストラクタ：一時ディレクトリの削除"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                logger.error(f"一時ディレクトリの削除中にエラーが発生しました: {e}")

    def _process_image(self, pdf: canvas.Canvas, img_path: str, row: int, col: int,
                      col_width_pt: float, row_height_pt: float, page_height: float,
                      temp_dir: str) -> None:
        """個々の画像を処理してPDFに配置する"""
        with Image.open(img_path) as img:
            # アスペクト比を維持したままセル内に収まるようリサイズ
            img_width, img_height = img.size
            img_aspect = img_width / img_height
            cell_aspect = col_width_pt / row_height_pt
            
            if img_aspect > cell_aspect:
                new_width = col_width_pt
                new_height = col_width_pt / img_aspect
            else:
                new_height = row_height_pt
                new_width = row_height_pt * img_aspect
            
            # セル内でセンタリング
            x_offset = col * col_width_pt + (col_width_pt - new_width) / 2
            y_offset = page_height - (row + 1) * row_height_pt + (row_height_pt - new_height) / 2
            
            img = img.resize((int(new_width), int(new_height)))
            
            # カラーモードの変換
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img_cmyk = img.convert('CMYK')
            
            temp_img_path = os.path.join(temp_dir, f"temp_{row}_{col}.jpg")
            img_cmyk.save(temp_img_path)
            
            pdf.drawImage(temp_img_path, x_offset, y_offset, new_width, new_height)

    def _draw_grid_lines(self, pdf: canvas.Canvas, cols: int, rows: int,
                        col_width_pt: float, row_height_pt: float,
                        page_width: float, page_height: float) -> None:
        """グリッド線を描画する"""
        r, g, b = (self.settings.grid_color.red() / 255.0,
                  self.settings.grid_color.green() / 255.0,
                  self.settings.grid_color.blue() / 255.0)
        pdf.setStrokeColorRGB(r, g, b)
        pdf.setLineWidth(self.settings.grid_width)
        
        # 垂直線
        for col in range(cols + 1):
            x = col * col_width_pt
            pdf.line(x, 0, x, page_height)
        
        # 水平線
        for row in range(rows + 1):
            y = page_height - row * row_height_pt
            pdf.line(0, y, page_width, y)


class ImageGridApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_paths: List[str] = []
        self.settings = GridSettings.load_from_file()  # 設定を読み込み
        self.preview_labels: List[QLabel] = []
        self.pdf_thread: Optional[PDFGenerationThread] = None
        self.progress_dialog: Optional[QProgressDialog] = None
        self.initUI()

    def initUI(self) -> None:
        """UIコンポーネントの初期化"""
        main_layout = QVBoxLayout()
        controls_layout = QVBoxLayout()

        # コントロールの初期化
        self._init_image_controls(controls_layout)
        self._init_grid_controls(controls_layout)
        self._init_preview_area(main_layout)

        main_layout.addLayout(controls_layout)
        self.setLayout(main_layout)
        self.setAcceptDrops(True)
        self.setWindowTitle("画像グリッド作成ツール")
        self.resize(600, 500)

        self.update_preview()

    def _init_image_controls(self, layout: QVBoxLayout) -> None:
        """画像関連のコントロールを初期化"""
        self.btn_add_images = QPushButton('画像を追加')
        self.btn_add_images.clicked.connect(self.load_images)
        layout.addWidget(self.btn_add_images)

    def _init_grid_controls(self, layout: QVBoxLayout) -> None:
        """グリッド設定関連のコントロールを初期化"""
        # 行の高さ設定
        self.row_height_spinbox = QDoubleSpinBox()
        self.row_height_spinbox.setRange(10.0, 297.0)
        self.row_height_spinbox.setValue(self.settings.row_height_mm)
        self.row_height_spinbox.setSuffix(" mm")
        self.row_height_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("行の高さ:"))
        layout.addWidget(self.row_height_spinbox)

        # 列の幅設定
        self.col_width_spinbox = QDoubleSpinBox()
        self.col_width_spinbox.setRange(10.0, 210.0)
        self.col_width_spinbox.setValue(self.settings.col_width_mm)
        self.col_width_spinbox.setSuffix(" mm")
        self.col_width_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("列の幅:"))
        layout.addWidget(self.col_width_spinbox)

        # グリッド線の設定
        self._init_grid_line_controls(layout)

        # ページサイズ選択
        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems(["A4", "A3"])
        # 保存された設定に基づいて初期選択を設定
        self.page_size_combo.setCurrentText("A4" if self.settings.page_size == A4 else "A3")
        self.page_size_combo.currentTextChanged.connect(self.update_page_size)
        layout.addWidget(QLabel("用紙サイズ:"))
        layout.addWidget(self.page_size_combo)

        # PDF生成ボタン
        self.btn_generate_pdf = QPushButton('PDFを作成')
        self.btn_generate_pdf.clicked.connect(self.generate_pdf)
        layout.addWidget(self.btn_generate_pdf)

    def _init_grid_line_controls(self, layout: QVBoxLayout) -> None:
        """グリッド線関連のコントロールを初期化"""
        self.grid_line_checkbox = QCheckBox("グリッド線を表示")
        self.grid_line_checkbox.setChecked(self.settings.grid_line_visible)
        self.grid_line_checkbox.stateChanged.connect(self.update_grid)
        layout.addWidget(self.grid_line_checkbox)
        
        self.grid_color_btn = QPushButton("グリッド線の色")
        self.grid_color_btn.clicked.connect(self.select_grid_color)
        layout.addWidget(self.grid_color_btn)
        
        self.grid_width_spinbox = QSpinBox()
        self.grid_width_spinbox.setRange(1, 5)
        self.grid_width_spinbox.setValue(self.settings.grid_width)
        self.grid_width_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("線の太さ:"))
        layout.addWidget(self.grid_width_spinbox)

    def _init_preview_area(self, layout: QVBoxLayout) -> None:
        """プレビューエリアを初期化"""
        self.preview_area_scroll = QScrollArea()
        self.preview_area_grid = QGridLayout()
        self.preview_area_grid.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_area_widget = QWidget()
        self.preview_area_widget.setLayout(self.preview_area_grid)
        self.preview_area_scroll.setWidget(self.preview_area_widget)
        self.preview_area_scroll.setWidgetResizable(True)
        layout.addWidget(self.preview_area_scroll)

    def load_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "画像を選択", "", "Images (*.png *.jpg *.jpeg)")
        if files:
            self.image_paths.extend(files)
            self.update_preview()

    def update_grid(self):
        self.settings.row_height_mm = self.row_height_spinbox.value()
        self.settings.col_width_mm = self.col_width_spinbox.value()
        self.settings.grid_line_visible = self.grid_line_checkbox.isChecked()
        self.settings.grid_width = self.grid_width_spinbox.value()
        self.update_preview()

    def update_page_size(self, size_text):
        if size_text == "A4":
            self.settings.page_size = A4
            self.row_height_spinbox.setRange(10.0, 297.0)  # A4の高さ制限
            self.col_width_spinbox.setRange(10.0, 210.0)   # A4の幅制限
        elif size_text == "A3":
            self.settings.page_size = A3
            self.row_height_spinbox.setRange(10.0, 420.0)  # A3の高さ制限
            self.col_width_spinbox.setRange(10.0, 297.0)   # A3の幅制限
        self.update_preview()

    @lru_cache(maxsize=100)
    def _create_thumbnail(self, img_path: str) -> QPixmap:
        """画像のサムネイルを生成（キャッシュ付き）"""
        # サムネイルを生成
        pixmap = QPixmap(img_path)
        thumbnail = pixmap.scaled(THUMBNAIL_SIZE[0], THUMBNAIL_SIZE[1],
                                Qt.AspectRatioMode.KeepAspectRatio,
                                Qt.TransformationMode.SmoothTransformation)
        return thumbnail

    def update_preview(self):
        # 既存のプレビューをクリア
        for label in self.preview_labels:
            label.clear()
            label.setParent(None)
        self.preview_labels = []

        # グリッドコンテナのクリア
        while self.preview_area_grid.count():
            item = self.preview_area_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.image_paths:
            return

        # mm単位をポイントに変換 (1mm = 2.83465pt)
        MM_TO_PT = 2.83465
        page_width, page_height = self.settings.page_size
        
        # プレビューのサイズを計算（A4/A3の比率を保持）
        preview_height = DEFAULT_PREVIEW_HEIGHT
        preview_width = int(preview_height * (page_width / page_height))
        
        # プレビュー用のフレームを作成（用紙を模したフレーム）
        self.preview_frame = QFrame()
        self.preview_frame.setFixedSize(preview_width, preview_height)
        self.preview_frame.setFrameShape(QFrame.Shape.Box)
        self.preview_frame.setStyleSheet("background-color: white;")
        
        # 行と列の数を計算
        col_width_pt = self.settings.col_width_mm * MM_TO_PT
        row_height_pt = self.settings.row_height_mm * MM_TO_PT
        cols = max(1, int(page_width / col_width_pt))
        rows = max(1, int(page_height / row_height_pt))
        
        # プレビューでのセルサイズを計算
        cell_width = preview_width / (page_width / col_width_pt)
        cell_height = preview_height / (page_height / row_height_pt)
        
        # 画像を描画するためのpaintEventを設定
        def paint_preview(event):
            painter = QPainter(self.preview_frame)
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
            
            # 画像の描画
            for row in range(rows):
                for col in range(cols):
                    cell_index = row * cols + col
                    img_index = cell_index % len(self.image_paths) if self.image_paths else 0
                    
                    if self.image_paths:
                        img_path = self.image_paths[img_index]
                        thumbnail = self._create_thumbnail(img_path)
                        
                        # セルのサイズとアスペクト比を計算
                        cell_rect_width = cell_width
                        cell_rect_height = cell_height
                        cell_aspect = cell_rect_width / cell_rect_height
                        
                        # 画像のアスペクト比を計算
                        img_aspect = thumbnail.width() / thumbnail.height()
                        
                        # アスペクト比に基づいてサイズを調整
                        if img_aspect > cell_aspect:
                            new_width = cell_rect_width
                            new_height = cell_rect_width / img_aspect
                        else:
                            new_height = cell_rect_height
                            new_width = cell_rect_height * img_aspect
                        
                        # セル内での位置を計算（センタリング）
                        x = col * cell_width + (cell_width - new_width) / 2
                        y = row * cell_height + (cell_height - new_height) / 2
                        
                        # 画像を描画
                        target_rect = QRectF(x, y, new_width, new_height)
                        source_rect = QRectF(thumbnail.rect())
                        painter.drawPixmap(target_rect, thumbnail, source_rect)
            
            # グリッド線の描画
            if self.settings.grid_line_visible:
                pen = QPen(self.settings.grid_color)
                pen.setWidth(self.settings.grid_width)
                painter.setPen(pen)
                
                # 垂直線
                for col in range(cols + 1):
                    x = col * cell_width
                    painter.drawLine(int(x), 0, int(x), preview_height)
                
                # 水平線
                for row in range(rows + 1):
                    y = row * cell_height
                    painter.drawLine(0, int(y), preview_width, int(y))
            
            painter.end()
        
        # paintEventを設定
        self.preview_frame.paintEvent = paint_preview
        
        # プレビューフレームをレイアウトに追加
        self.preview_area_grid.addWidget(self.preview_frame)
        self.preview_frame.update()

    def select_grid_color(self):
        """グリッド線の色を選択するダイアログを表示"""
        color = QColorDialog.getColor(self.settings.grid_color, self, "グリッド線の色を選択")
        if color.isValid():
            self.settings.grid_color = color
            self.update_preview()

    def generate_pdf(self):
        if not self.image_paths:
            QMessageBox.warning(self, "警告", "画像が追加されていません。")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "PDFを保存", "", "PDF Files (*.pdf)"
        )
        if not file_path:
            return  # ユーザーがキャンセルした場合
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        # 進捗ダイアログの作成
        self.progress_dialog = QProgressDialog("PDFを生成中...", "キャンセル", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setAutoReset(True)

        # PDF生成スレッドの作成と開始
        self.pdf_thread = PDFGenerationThread(
            self.image_paths,
            self.settings
        )
        
        # シグナルの接続
        self.pdf_thread.finished.connect(lambda temp_path, temp_dir: self.on_pdf_generation_finished(temp_path, temp_dir, file_path))
        self.pdf_thread.error.connect(self.on_pdf_generation_error)
        self.pdf_thread.progress.connect(self.progress_dialog.setValue)
        self.progress_dialog.canceled.connect(self.pdf_thread.terminate)
        
        # スレッドの開始
        self.pdf_thread.start()
        self.progress_dialog.show()

    def on_pdf_generation_finished(self, temp_path: str, temp_dir: str, final_path: str):
        """PDF生成完了時の処理"""
        try:
            # 一時ファイルを最終保存先にコピー
            import shutil
            shutil.copy2(temp_path, final_path)
            QMessageBox.information(self, "完了", f"PDFを作成しました: {final_path}")
        except Exception as e:
            logger.error(f"PDFの保存中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"PDFの保存中にエラーが発生しました: {e}")
        finally:
            # 一時ディレクトリの削除
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.error(f"一時ディレクトリの削除中にエラーが発生しました: {e}")

    def on_pdf_generation_error(self, error_message: str):
        """PDF生成エラー時の処理"""
        QMessageBox.critical(self, "エラー", f"PDFの生成中にエラーが発生しました: {error_message}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".png", ".jpg", ".jpeg")):
                self.image_paths.append(file_path)
        self.update_preview()
        print("画像を追加しました") # ←動作確認用

    def closeEvent(self, event: Any) -> None:
        """アプリケーション終了時の処理"""
        try:
            self.settings.save_to_file()  # 設定を保存
            self._create_thumbnail.cache_clear()  # サムネイルキャッシュをクリア
        except Exception as e:
            logger.error(f"設定の保存中にエラーが発生しました: {e}")
            QMessageBox.warning(self, "警告", "設定の保存に失敗しました。")
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageGridApp()
    ex.show()
    sys.exit(app.exec()) 