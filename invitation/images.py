"""Validate and re-encode an uploaded invitation image with Pillow.

The type is decided by reading the bytes, never by the file name or the
browser's MIME type. The output is re-encoded (WebP, or JPEG when this Pillow
has no WebP), rotated by EXIF and stripped of all metadata (GPS, camera), so
nothing hidden in the original reaches the public page.
"""
import io
import uuid

from PIL import Image, ImageOps, UnidentifiedImageError, features

MAX_BYTES = 8 * 1024 * 1024
MIN_SHORT_SIDE = 600
MAX_SIDE = 8000
MAX_PIXELS = 50_000_000
OUT_LONGEST = 2400
ALLOWED = {"JPEG", "PNG", "WEBP"}
KEY_RE = r"^[0-9a-f]{32}\.(webp|jpg)$"


class ImageError(ValueError):
    pass


def process(blob):
    """bytes -> (encoded bytes, info dict with key/mime/width/height)."""
    if not blob:
        raise ImageError("Choose an image file.")
    if len(blob) > MAX_BYTES:
        raise ImageError("The image is larger than 8 MB.")
    try:
        with Image.open(io.BytesIO(blob)) as probe:
            fmt = probe.format
            w, h = probe.size
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError):
        raise ImageError("Upload a JPG, PNG or WebP image.")
    if fmt not in ALLOWED:
        raise ImageError("Upload a JPG, PNG or WebP image.")
    if max(w, h) > MAX_SIDE or w * h > MAX_PIXELS:
        raise ImageError(f"The image is too large ({w}×{h}). Use at most {MAX_SIDE} px per side.")
    if min(w, h) < MIN_SHORT_SIDE:
        raise ImageError(f"The image is too small ({w}×{h}). Use at least {MIN_SHORT_SIDE} px on the short side.")
    try:
        with Image.open(io.BytesIO(blob)) as im:
            im.load()
            im = ImageOps.exif_transpose(im)
            has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
            im = im.convert("RGBA" if has_alpha else "RGB")
            im.thumbnail((OUT_LONGEST, OUT_LONGEST), Image.LANCZOS)
            out = io.BytesIO()
            if features.check("webp"):
                im.save(out, "WEBP", quality=85, method=4)
                ext, mime = "webp", "image/webp"
            else:  # pragma: no cover - every current Pillow wheel has WebP
                if im.mode == "RGBA":
                    im = im.convert("RGB")
                im.save(out, "JPEG", quality=88, optimize=True)
                ext, mime = "jpg", "image/jpeg"
            width, height = im.size
    except (OSError, ValueError, Image.DecompressionBombError):
        raise ImageError("The image could not be read. Try another file.")
    return out.getvalue(), {"key": f"{uuid.uuid4().hex}.{ext}", "mime": mime,
                            "width": width, "height": height}
