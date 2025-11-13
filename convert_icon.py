from PIL import Image
import os

def create_multi_size_ico():
    """从多个PNG文件创建多尺寸ICO"""
    sizes = [16, 32, 48, 64, 128]
    icon_images = []
    
    for size in sizes:
        png_file = f"icon_{size}.png"
        if os.path.exists(png_file):
            img = Image.open(png_file)
            if img.size != (size, size):
                img = img.resize((size, size), Image.Resampling.LANCZOS)
            icon_images.append(img)
            print(f"添加 {size}x{size} 图标")
        else:
            print(f"警告: 找不到 {png_file}")
    
    if icon_images:
        icon_images[0].save(
            "icon.ico",
            format='ICO',
            sizes=[img.size for img in icon_images],
            append_images=icon_images[1:]
        )
        print("成功创建多尺寸ICO文件: icon.ico")
    else:
        print("错误: 未找到任何图标文件")

if __name__ == "__main__":
    create_multi_size_ico()