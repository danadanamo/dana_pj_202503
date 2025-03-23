import logging
import os
import sys
import tempfile
from typing import List, Optional

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


class PDFGenerationThread(QThread):
    """PDF生成をバックグラウンドで実行するスレッド"""
    finished = pyqtSignal(str)  # 成功時にファイルパスを送信
    error = pyqtSignal(str)     # エラー時にメッセージを送信
    progress = pyqtSignal(int)  # 進捗状況を送信

    def __init__(self, image_paths: List[str], page_size: tuple, row_height_mm: float,
                 col_width_mm: float, grid_line_visible: bool, grid_color: QColor,
                 grid_width: int):
        super().__init__()
        self.image_paths = image_paths
        self.page_size = page_size
        self.row_height_mm = row_height_mm
        self.col_width_mm = col_width_mm
        self.grid_line_visible = grid_line_visible
        self.grid_color = grid_color
        self.grid_width = grid_width

    def run(self):
        try:
            # 一時ディレクトリの作成
            with tempfile.TemporaryDirectory() as temp_dir:
                file_path = os.path.join(temp_dir, "output.pdf")
                pdf = canvas.Canvas(file_path, pagesize=self.page_size)
                
                # mm単位をポイントに変換 (1mm = 2.83465pt)
                MM_TO_PT = 2.83465
                page_width, page_height = self.page_size
                
                # 行と列の数を計算
                col_width_pt = self.col_width_mm * MM_TO_PT
                row_height_pt = self.row_height_mm * MM_TO_PT
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
                                img_path = self.image_paths[img_index]
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
                                    
                                    # RGBAモードの画像をCMYKモードに変換
                                    if img.mode == 'RGBA':
                                        img = img.convert('RGB')
                                    
                                    # RGBをCMYKに変換
                                    img_cmyk = img.convert('CMYK')
                                    
                                    temp_img_path = os.path.join(temp_dir, f"temp_{row}_{col}.jpg")
                                    img_cmyk.save(temp_img_path)
                                    
                                    pdf.drawImage(temp_img_path, x_offset, y_offset, new_width, new_height)
                            except UnidentifiedImageError as e:
                                logger.error(f"画像の読み込みに失敗しました: {img_path}, エラー: {e}")
                                continue
                            except Exception as e:
                                logger.error(f"画像の処理中にエラーが発生しました: {img_path}, エラー: {e}")
                                continue
                        
                        processed_cells += 1
                        progress = int((processed_cells / total_cells) * 100)
                        self.progress.emit(progress)
                
                # グリッド線の描画
                if self.grid_line_visible:
                    r, g, b = self.grid_color.red() / 255.0, self.grid_color.green() / 255.0, self.grid_color.blue() / 255.0
                    pdf.setStrokeColorRGB(r, g, b)
                    pdf.setLineWidth(self.grid_width)
                    
                    # 垂直線
                    for col in range(cols + 1):
                        x = col * col_width_pt
                        pdf.line(x, 0, x, page_height)
                    
                    # 水平線
                    for row in range(rows + 1):
                        y = page_height - row * row_height_pt
                        pdf.line(0, y, page_width, y)
                
                pdf.save()
                self.finished.emit(file_path)
                
        except Exception as e:
            logger.error(f"PDF生成中にエラーが発生しました: {e}")
            self.error.emit(str(e))


class ImageGridApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_paths: List[str] = []
        # mm単位での行の高さと列の幅（デフォルト値）
        self.row_height_mm = 100.0
        self.col_width_mm = 100.0
        self.page_size = A4
        self.preview_labels: List[QLabel] = []
        # グリッド線の初期設定
        self.grid_line_visible = True
        self.grid_color = QColor(0, 0, 0)  # 黒
        self.grid_width = 1
        self.pdf_thread: Optional[PDFGenerationThread] = None
        self.progress_dialog: Optional[QProgressDialog] = None
        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()

        controls_layout = QVBoxLayout()

        # 画像追加ボタン
        self.btn_add_images = QPushButton('画像を追加')
        self.btn_add_images.clicked.connect(self.load_images)
        controls_layout.addWidget(self.btn_add_images)

        # 行の高さ設定 (mm単位)
        self.row_height_spinbox = QDoubleSpinBox()
        self.row_height_spinbox.setRange(10.0, 297.0)  # A4の高さ制限
        self.row_height_spinbox.setValue(self.row_height_mm)
        self.row_height_spinbox.setSuffix(" mm")
        self.row_height_spinbox.valueChanged.connect(self.update_grid)
        controls_layout.addWidget(QLabel("行の高さ:"))
        controls_layout.addWidget(self.row_height_spinbox)

        # 列の幅設定 (mm単位)
        self.col_width_spinbox = QDoubleSpinBox()
        self.col_width_spinbox.setRange(10.0, 210.0)  # A4の幅制限
        self.col_width_spinbox.setValue(self.col_width_mm)
        self.col_width_spinbox.setSuffix(" mm")
        self.col_width_spinbox.valueChanged.connect(self.update_grid)
        controls_layout.addWidget(QLabel("列の幅:"))
        controls_layout.addWidget(self.col_width_spinbox)

        # グリッド線の表示/非表示
        self.grid_line_checkbox = QCheckBox("グリッド線を表示")
        self.grid_line_checkbox.setChecked(True)
        self.grid_line_checkbox.stateChanged.connect(self.update_preview)
        controls_layout.addWidget(self.grid_line_checkbox)
        
        # グリッド線の色設定
        self.grid_color_btn = QPushButton("グリッド線の色")
        self.grid_color_btn.clicked.connect(self.select_grid_color)
        controls_layout.addWidget(self.grid_color_btn)
        
        # グリッド線の太さ
        self.grid_width_spinbox = QSpinBox()
        self.grid_width_spinbox.setRange(1, 5)
        self.grid_width_spinbox.setValue(1)
        self.grid_width_spinbox.valueChanged.connect(self.update_preview)
        controls_layout.addWidget(QLabel("線の太さ:"))
        controls_layout.addWidget(self.grid_width_spinbox)

        # ページサイズ選択
        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems(["A4", "A3"])
        self.page_size_combo.currentTextChanged.connect(self.update_page_size)
        controls_layout.addWidget(QLabel("用紙サイズ:"))
        controls_layout.addWidget(self.page_size_combo)

        # PDF生成ボタン
        self.btn_generate_pdf = QPushButton('PDFを作成')
        self.btn_generate_pdf.clicked.connect(self.generate_pdf)
        controls_layout.addWidget(self.btn_generate_pdf)

        main_layout.addLayout(controls_layout)

        # プレビューエリア
        self.preview_area_scroll = QScrollArea()
        self.preview_area_grid = QGridLayout()
        self.preview_area_grid.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_area_widget = QWidget()
        self.preview_area_widget.setLayout(self.preview_area_grid)
        self.preview_area_scroll.setWidget(self.preview_area_widget)
        self.preview_area_scroll.setWidgetResizable(True)
        main_layout.addWidget(self.preview_area_scroll)

        self.setLayout(main_layout)
        self.setAcceptDrops(True)
        self.setWindowTitle("画像グリッド作成ツール")
        self.resize(600, 500)

        self.update_preview()

    def load_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "画像を選択", "", "Images (*.png *.jpg *.jpeg)")
        if files:
            self.image_paths.extend(files)
            self.update_preview()

    def update_grid(self):
        self.row_height_mm = self.row_height_spinbox.value()
        self.col_width_mm = self.col_width_spinbox.value()
        self.grid_line_visible = self.grid_line_checkbox.isChecked()
        self.grid_width = self.grid_width_spinbox.value()
        self.update_preview()

    def update_page_size(self, size_text):
        if size_text == "A4":
            self.page_size = A4
            self.row_height_spinbox.setRange(10.0, 297.0)  # A4の高さ制限
            self.col_width_spinbox.setRange(10.0, 210.0)   # A4の幅制限
        elif size_text == "A3":
            self.page_size = A3
            self.row_height_spinbox.setRange(10.0, 420.0)  # A3の高さ制限
            self.col_width_spinbox.setRange(10.0, 297.0)   # A3の幅制限
        self.update_preview()

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
        page_width, page_height = self.page_size
        
        # プレビューのサイズを計算（A4/A3の比率を保持）
        preview_height = 600  # プレビューの高さを固定
        preview_width = int(preview_height * (page_width / page_height))
        
        # プレビュー用のフレームを作成（用紙を模したフレーム）
        self.preview_frame = QFrame()
        self.preview_frame.setFixedSize(preview_width, preview_height)
        self.preview_frame.setFrameShape(QFrame.Shape.Box)
        self.preview_frame.setStyleSheet("background-color: white;")
        
        # 行と列の数を計算
        col_width_pt = self.col_width_mm * MM_TO_PT
        row_height_pt = self.row_height_mm * MM_TO_PT
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
                        pixmap = QPixmap(img_path)
                        
                        # セルのサイズとアスペクト比を計算
                        cell_rect_width = cell_width
                        cell_rect_height = cell_height
                        cell_aspect = cell_rect_width / cell_rect_height
                        
                        # 画像のアスペクト比を計算
                        img_aspect = pixmap.width() / pixmap.height()
                        
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
                        painter.drawPixmap(target_rect, pixmap, pixmap.rect())
            
            # グリッド線の描画
            if self.grid_line_visible:
                pen = QPen(self.grid_color)
                pen.setWidth(self.grid_width)
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
        color = QColorDialog.getColor(self.grid_color, self, "グリッド線の色を選択")
        if color.isValid():
            self.grid_color = color
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
            self.page_size,
            self.row_height_mm,
            self.col_width_mm,
            self.grid_line_visible,
            self.grid_color,
            self.grid_width
        )
        
        # シグナルの接続
        self.pdf_thread.finished.connect(lambda path: self.on_pdf_generation_finished(path, file_path))
        self.pdf_thread.error.connect(self.on_pdf_generation_error)
        self.pdf_thread.progress.connect(self.progress_dialog.setValue)
        self.progress_dialog.canceled.connect(self.pdf_thread.terminate)
        
        # スレッドの開始
        self.pdf_thread.start()
        self.progress_dialog.show()

    def on_pdf_generation_finished(self, temp_path: str, final_path: str):
        """PDF生成完了時の処理"""
        try:
            # 一時ファイルを最終保存先にコピー
            import shutil
            shutil.copy2(temp_path, final_path)
            QMessageBox.information(self, "完了", f"PDFを作成しました: {final_path}")
        except Exception as e:
            logger.error(f"PDFの保存中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"PDFの保存中にエラーが発生しました: {e}")

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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageGridApp()
    ex.show()
    sys.exit(app.exec()) 