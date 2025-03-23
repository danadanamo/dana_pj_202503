import os
import sys
import tempfile

from PIL import Image
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QImage, QPixmap
from PyQt6.QtWidgets import (QApplication, QComboBox, QDoubleSpinBox,
                             QFileDialog, QGridLayout, QLabel, QMessageBox,
                             QPushButton, QScrollArea, QSpinBox, QVBoxLayout,
                             QWidget)
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
        for label in self.preview_labels:
            label.clear()
            label.setParent(None)
        self.preview_labels = []

        if not self.image_paths:
            return

        # mm単位をポイントに変換 (1mm = 2.83465pt)
        MM_TO_PT = 2.83465
        page_width, page_height = self.page_size
        
        # 行と列の数を計算
        col_width_pt = self.col_width_mm * MM_TO_PT
        row_height_pt = self.row_height_mm * MM_TO_PT
        cols = max(1, int(page_width / col_width_pt))
        rows = max(1, int(page_height / row_height_pt))
        
        # プレビュー用のグリッドに画像を配置
        for i, img_path in enumerate(self.image_paths):
            if i >= rows * cols:
                break  # プレビューは1ページだけ
                
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            pixmap = QPixmap(img_path)

            max_preview_width = 150
            max_preview_height = 150
            pixmap_scaled = pixmap.scaled(max_preview_width, max_preview_height, 
                                         Qt.AspectRatioMode.KeepAspectRatio, 
                                         Qt.TransformationMode.SmoothTransformation)

            label.setPixmap(pixmap_scaled)
            self.preview_labels.append(label)
            row = i // cols
            col = i % cols
            self.preview_area_grid.addWidget(label, row, col)

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