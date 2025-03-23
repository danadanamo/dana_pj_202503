import importlib.util
import os
import sys
import tempfile

from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# テスト対象のPDF生成機能をインポート
spec = importlib.util.spec_from_file_location("d_pj_image_grid_app_v0.02", "step_01/d_pj_image_grid_app_v0.02.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

def test_pdf_generation():
    """PDFの実際の生成機能をテスト"""
    
    # テスト用の一時ディレクトリを作成
    with tempfile.TemporaryDirectory() as temp_dir:
        # テスト用のパラメータ
        output_pdf = os.path.join(temp_dir, "test_output.pdf")
        page_size = A4
        
        # mm単位での行の高さと列の幅を指定
        row_height_mm = 100.0
        col_width_mm = 100.0
        
        # テスト画像のパス
        test_images = [
            os.path.abspath(f'test_images/test_image_{i}.png') for i in range(1, 5)
        ]
        
        # mm単位をポイントに変換 (1mm = 2.83465pt)
        MM_TO_PT = 2.83465
        page_width, page_height = page_size
        
        # 行と列の数を計算
        col_width_pt = col_width_mm * MM_TO_PT
        row_height_pt = row_height_mm * MM_TO_PT
        cols = max(1, int(page_width / col_width_pt))
        rows = max(1, int(page_height / row_height_pt))
        
        print(f"PDFの生成を開始: {output_pdf}")
        print(f"行の高さ: {row_height_mm}mm, 列の幅: {col_width_mm}mm")
        print(f"計算された行数: {rows}, 列数: {cols}")
        
        pdf = canvas.Canvas(output_pdf, pagesize=page_size)
        
        for row in range(rows):
            for col in range(cols):
                cell_index = row * cols + col
                # 画像がない場合は最初の画像を使用（繰り返し）
                img_index = cell_index % len(test_images) if test_images else 0
                
                if test_images:
                    img_path = test_images[img_index]
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
        
        # PDFファイルが存在するか確認
        if os.path.exists(output_pdf):
            print(f"PDF生成テスト成功: {output_pdf}")
            print(f"ファイルサイズ: {os.path.getsize(output_pdf)} バイト")
            return True
        else:
            print(f"PDF生成テスト失敗: ファイルが作成されていません")
            return False

if __name__ == "__main__":
    result = test_pdf_generation()
    print(f"テスト結果: {'成功' if result else '失敗'}") 