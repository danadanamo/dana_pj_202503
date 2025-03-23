import os
import sys
import tempfile

from PIL import Image
from PyQt6.QtCore import QRectF, QSize, Qt
from PyQt6.QtGui import (QColor, QDragEnterEvent, QDropEvent, QImage, QPainter,
                         QPen, QPixmap)
from PyQt6.QtWidgets import (QApplication, QCheckBox, QColorDialog, QComboBox,
                             QDoubleSpinBox, QFileDialog, QFrame, QGridLayout,
                             QLabel, QMessageBox, QPushButton, QScrollArea,
                             QSpinBox, QVBoxLayout, QWidget)
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas


class ImageGridApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_paths = []
        # mm単位での行の高さと列の幅（デフォルト値）
        self.row_height_mm = 100.0
        self.col_width_mm = 100.0
        self.page_size = A4
        self.preview_labels = []
        # グリッド線の初期設定
        self.grid_line_visible = True
        self.grid_color = QColor(0, 0, 0)  # 黒
        self.grid_width = 1
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

        # mm単位をポイントに変換 (1mm = 2.83465pt)
        MM_TO_PT = 2.83465
        page_width, page_height = self.page_size
        
        # 行と列の数を計算
        col_width_pt = self.col_width_mm * MM_TO_PT
        row_height_pt = self.row_height_mm * MM_TO_PT
        cols = max(1, int(page_width / col_width_pt))
        rows = max(1, int(page_height / row_height_pt))

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf = canvas.Canvas(file_path, pagesize=self.page_size)
            
            # 最初に画像を配置
            for row in range(rows):
                for col in range(cols):
                    cell_index = row * cols + col
                    # 画像がない場合は最初の画像を使用（繰り返し）
                    img_index = cell_index % len(self.image_paths) if self.image_paths else 0
                    
                    if self.image_paths:
                        img_path = self.image_paths[img_index]
                        img = Image.open(img_path)
                        
                        # アスペクト比を維持したままセル内に収まるようリサイズ
                        img_width, img_height = img.size
                        img_aspect = img_width / img_height
                        cell_aspect = col_width_pt / row_height_pt
                        
                        if img_aspect > cell_aspect:
                            # 画像が横長の場合
                            new_width = col_width_pt
                            new_height = col_width_pt / img_aspect
                        else:
                            # 画像が縦長の場合
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
            
            # 画像の配置後に罫線を描画
            if self.grid_line_visible:
                # RGB値を0-1の範囲に変換
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
        QMessageBox.information(self, "完了", f"PDFを作成しました: {file_path}")

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