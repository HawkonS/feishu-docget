import math
import os


from PIL import Image, ImageChops


def calculate_center_crop(source_width, source_height, frame_width, frame_height, tolerance=0.001):
    """Return normalized crop margins needed to fit an image into a frame.

    Feishu's public image-block response exposes the visible frame dimensions,
    but not the crop offsets. A differing aspect ratio therefore means the
    image was cropped; the best representation available through OpenAPI is a
    centered crop with the same visible aspect ratio.
    """
    try:
        source_width = float(source_width)
        source_height = float(source_height)
        frame_width = float(frame_width)
        frame_height = float(frame_height)
    except (TypeError, ValueError):
        return None

    dimensions = (source_width, source_height, frame_width, frame_height)
    if any(not math.isfinite(value) or value <= 0 for value in dimensions):
        return None

    source_ratio = source_width / source_height
    frame_ratio = frame_width / frame_height
    if abs(source_ratio - frame_ratio) / source_ratio <= tolerance:
        return None

    crop = {'left': 0.0, 'top': 0.0, 'right': 0.0, 'bottom': 0.0}
    if source_ratio > frame_ratio:
        horizontal_crop = (1.0 - frame_ratio / source_ratio) / 2.0
        crop['left'] = horizontal_crop
        crop['right'] = horizontal_crop
    else:
        vertical_crop = (1.0 - source_ratio / frame_ratio) / 2.0
        crop['top'] = vertical_crop
        crop['bottom'] = vertical_crop
    return crop


def get_image_center_crop(image_path, frame_width, frame_height, tolerance=0.001):
    """Calculate crop margins from a raster file and Feishu's frame size."""
    if not os.path.exists(image_path):
        return None
    try:
        with Image.open(image_path) as image:
            source_width, source_height = image.size
        return calculate_center_crop(
            source_width,
            source_height,
            frame_width,
            frame_height,
            tolerance=tolerance,
        )
    except Exception:
        return None


def get_image_dimensions(image_path):
    """Return raster dimensions as ``(width, height)`` when available."""
    if not os.path.exists(image_path):
        return None
    try:
        with Image.open(image_path) as image:
            return image.size
    except Exception:
        return None


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
