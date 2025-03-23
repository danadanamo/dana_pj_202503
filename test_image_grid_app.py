import os
import sys
import unittest

from PyQt6.QtCore import Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

from step_03.d_pj_image_grid_app_v004 import GridSettings, ImageGridApp


class TestImageGridApp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication(sys.argv)
        cls.window = ImageGridApp()

    def setUp(self):
        # 各テストの前にウィンドウをリセット
        self.window.image_paths = []
        self.window.update_preview()

    def test_01_load_images(self):
        """画像の読み込みテスト"""
        # テスト画像のパスを設定
        test_images = [
            'test_images/test1.png',
            'test_images/test2.jpg',
            'test_images/test3.jpeg'
        ]

        # 画像を追加
        for img_path in test_images:
            self.window.image_paths.append(img_path)
            self.window.update_preview()

        # 画像が正しく追加されたか確認
        self.assertEqual(len(self.window.image_paths), len(test_images))

    def test_02_grid_settings(self):
        """グリッド設定のテスト"""
        # 行の高さを変更
        self.window.row_height_spinbox.setValue(150.0)
        self.window.update_grid()
        self.assertEqual(self.window.row_height_spinbox.value(), 150.0)

        # 列の幅を変更
        self.window.col_width_spinbox.setValue(150.0)
        self.window.update_grid()
        self.assertEqual(self.window.col_width_spinbox.value(), 150.0)

    def test_03_page_settings(self):
        """ページ設定のテスト"""
        # A4サイズのテスト
        self.window.page_size_combo.setCurrentText("A4")
        self.window.update_page_size("A4")
        self.assertEqual(self.window.page_size_combo.currentText(), "A4")

        # A3サイズのテスト
        self.window.page_size_combo.setCurrentText("A3")
        self.window.update_page_size("A3")
        self.assertEqual(self.window.page_size_combo.currentText(), "A3")

    def test_04_settings_save_load(self):
        """設定の保存と読み込みのテスト"""
        # 設定を変更
        settings = GridSettings()
        settings.row_height_mm = 200.0
        settings.col_width_mm = 200.0
        settings.grid_line_visible = False

        # 設定を保存
        settings.save_to_file()

        # 新しい設定オブジェクトを作成して読み込み
        loaded_settings = GridSettings.load_from_file()

        # 設定が正しく保存・読み込みされたか確認
        self.assertEqual(loaded_settings.row_height_mm, 200.0)
        self.assertEqual(loaded_settings.col_width_mm, 200.0)
        self.assertEqual(loaded_settings.grid_line_visible, False)

    def test_05_error_handling(self):
        """エラーハンドリングのテスト"""
        # 存在しない画像ファイルを追加
        self.window.image_paths.append('nonexistent.png')
        self.window.update_preview()

        # 無効な設定ファイルのテスト
        with open('grid_settings.json', 'w') as f:
            f.write('invalid json')
        
        # 設定を読み込もうとする
        settings = GridSettings.load_from_file()
        self.assertIsNotNone(settings)

    @classmethod
    def tearDownClass(cls):
        # テスト用ファイルの削除
        if os.path.exists('grid_settings.json'):
            os.remove('grid_settings.json')
        if os.path.exists('grid_settings.json.backup'):
            os.remove('grid_settings.json.backup')
        cls.window.close()
        cls.app.quit()

if __name__ == '__main__':
    unittest.main() 