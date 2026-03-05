from PIL import Image, ImageChops
import os

def smart_crop(image_path, padding=50):
    if not os.path.exists(image_path):
        return False
    try:
        img = Image.open(image_path)
        if img.mode == 'RGBA':
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            rgb_img = background
        else:
            rgb_img = img.convert('RGB')
        bg = Image.new('RGB', rgb_img.size, (255, 255, 255))
        diff = ImageChops.difference(rgb_img, bg)
        bbox = diff.getbbox()
        if bbox:
            left, top, right, bottom = bbox
            width, height = rgb_img.size
            left = max(0, left - padding)
            top = max(0, top - padding)
            right = min(width, right + padding)
            bottom = min(height, bottom + padding)
            if left > 0 or top > 0 or right < width or (bottom < height):
                cropped = img.crop((left, top, right, bottom))
                cropped.save(image_path)
                return True
    except Exception as e:
        pass
    return False