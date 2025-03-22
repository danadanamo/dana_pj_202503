import os
import sys

from PIL import Image
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QImage, QPixmap
from PyQt6.QtWidgets import (QApplication, QFileDialog, QGridLayout, QLabel,
                             QPushButton, QScrollArea, QSpinBox, QVBoxLayout,
                             QWidget)
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas


class ImageGridApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_paths = []
        self.grid_rows = 2
        self.grid_cols = 2
        self.page_size = A4
        self.preview_labels = [] # プレビュー画像のQLabelを保持するリスト
        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()

        # 画像追加、グリッド設定、PDF出力ボタンなどを配置するレイアウト
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

        # PDF生成ボタン
        self.btn_generate_pdf = QPushButton('PDFを作成')
        self.btn_generate_pdf.clicked.connect(self.generate_pdf)
        controls_layout.addWidget(self.btn_generate_pdf)

        main_layout.addLayout(controls_layout)

        # プレビューエリア (スクロール可能にする)
        self.preview_area_scroll = QScrollArea()
        self.preview_area_grid = QGridLayout()
        self.preview_area_grid.setAlignment(Qt.AlignmentFlag.AlignCenter) # グリッドを中央寄せ
        self.preview_area_widget = QWidget() # グリッドレイアウトをセットするQWidget
        self.preview_area_widget.setLayout(self.preview_area_grid)
        self.preview_area_scroll.setWidget(self.preview_area_widget)
        self.preview_area_scroll.setWidgetResizable(True) # リサイズ可能にする
        main_layout.addWidget(self.preview_area_scroll)

        self.setLayout(main_layout)
        self.setAcceptDrops(True)
        self.setWindowTitle("画像グリッド作成ツール")
        self.resize(600, 500) # ウィンドウサイズを少し大きく

        self.update_preview() # 初期プレビュー表示

    def load_images(self):
        files, _ = QFileDialog.getOpenFileNames(self, "画像を選択", "", "Images (*.png *.jpg *.jpeg)")
        if files:
            self.image_paths.extend(files)
            self.update_preview() # 画像追加後にプレビュー更新

    def update_grid(self):
        self.grid_rows = self.row_spinbox.value()
        self.grid_cols = self.col_spinbox.value()
        self.update_preview() # グリッド数変更後にプレビュー更新

    def update_preview(self):
        # プレビューエリアをクリア
        for label in self.preview_labels:
            label.clear() # QLabelの内容をクリア
            label.setParent(None) # QLabelをレイアウトから削除
        self.preview_labels = [] # リストを空にする

        num_images = len(self.image_paths)
        grid_size = self.grid_rows * self.grid_cols

        for i in range(min(num_images, grid_size)): # グリッドサイズまたは画像枚数の少ない方まで表示
            img_path = self.image_paths[i]
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignCenter) # 画像を中央寄せ
            pixmap = QPixmap(img_path)

            # プレビュー画像の最大サイズ (仮の値。調整が必要かもしれません)
            max_preview_width = 200
            max_preview_height = 200
            pixmap_scaled = pixmap.scaled(max_preview_width, max_preview_height, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation) # サイズ調整

            label.setPixmap(pixmap_scaled)
            self.preview_labels.append(label) # QLabelをリストに追加
            row = i // self.grid_cols
            col = i % self.grid_cols
            self.preview_area_grid.addWidget(label, row, col) # グリッドレイアウトに追加


    def generate_pdf(self):
        if not self.image_paths:
            return

        page_width, page_height = self.page_size
        grid_size = (self.grid_cols, self.grid_rows)
        pdf = canvas.Canvas("output.pdf", pagesize=self.page_size)
        img_size = (page_width // grid_size[0], page_height // grid_size[1])

        x_offset, y_offset = 0, page_height - img_size[1]
        for idx, img_path in enumerate(self.image_paths[:grid_size[0] * grid_size[1]]):
            img = Image.open(img_path)
            img = img.resize((int(img_size[0]), int(img_size[1])))
            img.save("temp.jpg")
            pdf.drawImage("temp.jpg", x_offset, y_offset, img_size[0], img_size[1])

            x_offset += img_size[0]
            if (idx + 1) % grid_size[0] == 0:
                x_offset = 0
                y_offset -= img_size[1]

        pdf.save()
        print("PDFを作成しました: output.pdf")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".png", ".jpg", ".jpeg")):
                self.image_paths.append(file_path)
        self.update_preview() # ドラッグ＆ドロップ後にもプレビュー更新
        print("画像を追加しました")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageGridApp()
    ex.show()
    sys.exit(app.exec())