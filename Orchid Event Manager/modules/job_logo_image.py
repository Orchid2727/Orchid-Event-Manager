from __future__ import annotations

from collections import deque

from PIL import Image as PILImage


def remove_outer_near_white_background(
    image: PILImage.Image,
    *,
    threshold: int = 232,
    neutral_tolerance: int = 28,
) -> tuple[PILImage.Image, bool]:
    """Make only edge-connected pale canvas pixels transparent and trim it.

    The original upload is never changed.  This intentionally does not remove
    white enclosed by a logo outline, such as white artwork inside a city seal.
    """
    result = image.convert("RGBA")
    width, height = result.size
    if width < 1 or height < 1:
        return result, False

    pixels = result.load()
    seen = bytearray(width * height)
    pending: deque[tuple[int, int]] = deque()

    def is_outer_canvas(x: int, y: int) -> bool:
        red, green, blue, alpha = pixels[x, y]
        # Exported JPGs and screenshots often leave a very light gray, cream,
        # or blue-tinted halo around a white canvas.  Treat light *neutral*
        # pixels as canvas too, but only when they connect to the outer edge.
        # That preserves white regions enclosed by the dark ring of a seal.
        channels = (red, green, blue)
        return (
            alpha > 0
            and min(channels) >= threshold
            and max(channels) - min(channels) <= neutral_tolerance
        )

    def add_if_canvas(x: int, y: int) -> None:
        index = y * width + x
        if seen[index] or not is_outer_canvas(x, y):
            return
        seen[index] = 1
        pending.append((x, y))

    for x in range(width):
        add_if_canvas(x, 0)
        add_if_canvas(x, height - 1)
    for y in range(1, height - 1):
        add_if_canvas(0, y)
        add_if_canvas(width - 1, y)

    removed = 0
    while pending:
        x, y = pending.popleft()
        red, green, blue, _alpha = pixels[x, y]
        pixels[x, y] = (red, green, blue, 0)
        removed += 1
        if x:
            add_if_canvas(x - 1, y)
        if x + 1 < width:
            add_if_canvas(x + 1, y)
        if y:
            add_if_canvas(x, y - 1)
        if y + 1 < height:
            add_if_canvas(x, y + 1)

    # After the outer canvas is transparent, trim it from the in-memory copy.
    # This removes a remaining pale border or otherwise-empty rectangle from
    # the Current Event preview and lets the visible logo use its true bounds.
    alpha_box = result.getchannel("A").getbbox()
    trimmed = False
    if alpha_box and alpha_box != (0, 0, width, height):
        result = result.crop(alpha_box)
        trimmed = True

    return result, bool(removed or trimmed)
