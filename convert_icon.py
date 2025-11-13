from PIL import Image
import os
import sys

def main():
    if os.path.exists('icon.png'):
        print("Converting icon.png to icon.ico...")
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]
        img = Image.open('icon.png')
        icon_images = []
        
        for size in sizes:
            resized_img = img.resize(size, Image.Resampling.LANCZOS)
            icon_images.append(resized_img)
        
        icon_images[0].save(
            'icon.ico',
            format='ICO',
            sizes=[(img.width, img.height) for img in icon_images],
            append_images=icon_images[1:]
        )
        print("✅ Created icon.ico from icon.png")
    else:
        print("⚠️ No icon.png found, building without custom icon")

if __name__ == "__main__":
    main()