import sys
import os
import tempfile
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QGridLayout,
    QSpinBox, QScrollArea, QMessageBox, QComboBox
)
from PyQt6.QtGui import QPixmap, QDragEnterEvent, QDropEvent, QImage
from PyQt6.QtCore import Qt, QSize
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A3

class ImageGridApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_paths = []
        self.grid_rows = 2
        self.grid_cols = 2
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

        # グリッド行数設定
        self.row_spinbox = QSpinBox()
        self.row_spinbox.setRange(1, 10)
        self.row_spinbox.setValue(self.grid_rows)
        self.row_spinbox.valueChanged.connect(self.update_grid)
        controls_layout.addWidget(QLabel("行数:"))
        controls_layout.addWidget(self.row_spinbox)

        # グリッド列数設定
        self.col_spinbox = QSpinBox()
        self.col_spinbox.setRange(1, 10)
        self.col_spinbox.setValue(self.grid_cols)
        self.col_spinbox.valueChanged.connect(self.update_grid)
        controls_layout.addWidget(QLabel("列数:"))
        controls_layout.addWidget(self.col_spinbox)

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
        self.grid_rows = self.row_spinbox.value()
        self.grid_cols = self.col_spinbox.value()
        self.update_preview()

    def update_page_size(self, size_text):
        if size_text == "A4":
            self.page_size = A4
        elif size_text == "A3":
            self.page_size = A3
        self.update_preview()

    def update_preview(self):
        for label in self.preview_labels:
            label.clear()
            label.setParent(None)
        self.preview_labels = []

        num_images = len(self.image_paths)
        grid_size = self.grid_rows * self.grid_cols

        for i in range(min(num_images, grid_size)):
            img_path = self.image_paths[i]
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            pixmap = QPixmap(img_path)

            max_preview_width = 200
            max_preview_height = 200
            pixmap_scaled = pixmap.scaled(max_preview_width, max_preview_height, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

            label.setPixmap(pixmap_scaled)
            self.preview_labels.append(label)
            row = i // self.grid_cols
            col = i % self.grid_cols
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

        page_width, page_height = self.page_size
        grid_size = (self.grid_cols, self.grid_rows)

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf = canvas.Canvas(file_path, pagesize=self.page_size)
            img_size = (page_width // grid_size[0], page_height // grid_size[1])
            x_offset, y_offset = 0, page_height - img_size[1]

            for idx, img_path in enumerate(self.image_paths[:grid_size[0] * grid_size[1]]):
                img = Image.open(img_path)
                img = img.resize((int(img_size[0]), int(img_size[1])))
                temp_img_path = os.path.join(temp_dir, f"temp_{idx}.jpg")
                img.save(temp_img_path)
                pdf.drawImage(temp_img_path, x_offset, y_offset, img_size[0], img_size[1])

                x_offset += img_size[0]
                if (idx + 1) % grid_size[0] == 0:
                    x_offset = 0
                    y_offset -= img_size[1]

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