import os

from PIL import Image, ImageDraw


def create_test_image(size, color, filename):
    img = Image.new('RGB', size, color)
    draw = ImageDraw.Draw(img)
    # 画像にテキストを追加
    draw.text((10, 10), filename, fill='white')
    img.save(filename)

# テスト用ディレクトリの作成
os.makedirs('test_images', exist_ok=True)

# テスト画像の生成
create_test_image((800, 600), 'red', 'test_images/test1.png')
create_test_image((800, 600), 'blue', 'test_images/test2.jpg')
create_test_image((800, 600), 'green', 'test_images/test3.jpeg')

print("テスト画像が生成されました。") 