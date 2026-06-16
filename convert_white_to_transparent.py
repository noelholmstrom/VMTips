#!/usr/bin/env python3
"""
Convert near-white pixels in vmbild.jpg to transparent and save as vmbild_alpha.png
Usage: python convert_white_to_transparent.py [source] [dest] [threshold]
threshold default 250 (0-255, higher -> only very white become transparent)
"""
from PIL import Image
import sys, os

src = sys.argv[1] if len(sys.argv) > 1 else 'vmbild.jpg'
dst = sys.argv[2] if len(sys.argv) > 2 else 'vmbild_alpha.png'
try:
    threshold = int(sys.argv[3]) if len(sys.argv) > 3 else 250
except:
    threshold = 250

if not os.path.exists(src):
    print(f"Source not found: {src}")
    sys.exit(1)

im = Image.open(src).convert('RGBA')
px = im.load()
width, height = im.size
count_transparent = 0
for y in range(height):
    for x in range(width):
        r,g,b,a = px[x,y]
        if r >= threshold and g >= threshold and b >= threshold:
            px[x,y] = (r,g,b,0)
            count_transparent += 1

im.save(dst)
print(f"Saved {dst} ({count_transparent} pixels made transparent)")
