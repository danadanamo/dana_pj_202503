from PIL import Image, ImageDraw

# 異なる色のテスト画像を4つ作成
colors = ['red', 'green', 'blue', 'yellow']
size = (300, 300)

for i, color in enumerate(colors):
    img = Image.new('RGB', size, (255, 255, 255))
    draw = ImageDraw.Draw(img)
    if color == 'red':
        draw.rectangle([(50, 50), (250, 250)], fill=(255, 0, 0))
    elif color == 'green':
        draw.rectangle([(50, 50), (250, 250)], fill=(0, 255, 0))
    elif color == 'blue':
        draw.rectangle([(50, 50), (250, 250)], fill=(0, 0, 255))
    elif color == 'yellow':
        draw.rectangle([(50, 50), (250, 250)], fill=(255, 255, 0))
    
    img.save(f'test_images/test_image_{i+1}.png')
    
print("テスト画像が作成されました。") 