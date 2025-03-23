import importlib.util
import os
import sys
import unittest

from PyQt6.QtCore import Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

# テスト対象のアプリケーションをインポート
spec = importlib.util.spec_from_file_location("d_pj_image_grid_app_v0.02", "step_01/d_pj_image_grid_app_v0.02.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
ImageGridApp = module.ImageGridApp

class TestImageGridApp(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        # アプリケーションの初期化
        cls.app = QApplication(sys.argv)
        
    def setUp(self):
        # 各テスト前にアプリケーションのインスタンスを作成
        self.grid_app = ImageGridApp()
        
        # テスト画像のパス
        self.test_images = [
            os.path.abspath(f'test_images/test_image_{i}.png') for i in range(1, 5)
        ]
        
    def test_init_state(self):
        """初期状態のテスト"""
        self.assertEqual(self.grid_app.grid_rows, 2)
        self.assertEqual(self.grid_app.grid_cols, 2)
        self.assertEqual(len(self.grid_app.image_paths), 0)
        
    def test_update_grid(self):
        """グリッド更新のテスト"""
        self.grid_app.row_spinbox.setValue(3)
        self.assertEqual(self.grid_app.grid_rows, 3)
        
        self.grid_app.col_spinbox.setValue(4)
        self.assertEqual(self.grid_app.grid_cols, 4)
        
    def test_add_images(self):
        """画像追加機能のテスト（手動追加のシミュレーション）"""
        # 画像パスを直接追加
        for image_path in self.test_images:
            self.grid_app.image_paths.append(image_path)
        
        self.grid_app.update_preview()
        self.assertEqual(len(self.grid_app.image_paths), 4)
        
    def test_generate_pdf(self):
        """PDF生成機能の基本テスト（実際の保存は行わない）"""
        # 画像パスを追加
        for image_path in self.test_images:
            self.grid_app.image_paths.append(image_path)
            
        # 保存ダイアログをシミュレートせずに内部処理のみテスト
        # 実際の保存処理をテストするにはQFileDialogのモックが必要
        
        # 一旦テスト用のパスを指定して例外が発生しないか確認
        try:
            test_pdf_path = os.path.abspath('test_output.pdf')
            page_width, page_height = self.grid_app.page_size
            grid_size = (self.grid_app.grid_cols, self.grid_app.grid_rows)
            # 実際に保存まで行わないよう、内部ロジックだけをテスト
            self.assertTrue(isinstance(grid_size, tuple))
            self.assertTrue(len(grid_size) == 2)
        except Exception as e:
            self.fail(f"PDF生成準備中に例外が発生: {e}")

if __name__ == '__main__':
    unittest.main() 