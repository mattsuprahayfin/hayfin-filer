#!/usr/bin/env python3
"""
Generates simple placeholder PNG icons for the Hayfin Inbox Filer add-in.
Run once: python3 generate_icons.py
Requires: pip install pillow
"""
import os

sizes = [16, 32, 64, 80, 128]
os.makedirs('assets', exist_ok=True)

try:
    from PIL import Image, ImageDraw, ImageFont

    for size in sizes:
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Background circle
        margin = max(1, size // 10)
        draw.ellipse([margin, margin, size - margin, size - margin],
                     fill=(15, 15, 15), outline=(200, 255, 0), width=max(1, size // 16))

        # "F" letter
        font_size = max(8, size // 2)
        try:
            font = ImageFont.truetype('/System/Library/Fonts/Helvetica.ttc', font_size)
        except:
            font = ImageFont.load_default()

        text = 'F'
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (size - tw) // 2 - bbox[0]
        y = (size - th) // 2 - bbox[1]
        draw.text((x, y), text, fill=(200, 255, 0), font=font)

        img.save(f'assets/icon-{size}.png')
        print(f'Generated assets/icon-{size}.png')

    print('\nAll icons generated successfully.')

except ImportError:
    print('Pillow not installed. Install with: pip install pillow')
    print('Or replace assets/icon-*.png with your own 16x16, 32x32, 64x64, 80x80, 128x128 PNG files.')
