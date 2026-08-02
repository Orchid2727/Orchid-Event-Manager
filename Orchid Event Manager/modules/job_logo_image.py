from __future__ import annotations

from collections import deque
from os import PathLike

from PIL import Image as PILImage, ImageOps


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


def _rgb_distance(first: tuple[int, int, int], second: tuple[int, int, int]) -> float:
    """Return a lightweight RGB distance without adding another dependency."""
    return sum((left - right) ** 2 for left, right in zip(first, second)) ** 0.5


def remove_outer_edge_background(
    image: PILImage.Image,
    *,
    minimum_edge_coverage: float = 0.46,
) -> tuple[PILImage.Image, bool]:
    """Remove a single, edge-connected colored artwork background when safe.

    Many artwork references arrive as a JPG or screenshot on a light gray,
    cream, blue, or other solid-colored canvas.  The canvas should not appear
    as a rectangle on the outsourced report.  This routine identifies the
    dominant *outer-edge* color and removes only pixels connected to that
    edge.  It deliberately leaves enclosed areas alone, so white in a patch,
    seal, or logo stays intact.

    It opts out unless one color clearly owns most of at least three image
    edges.  That protects artwork whose actual design reaches the crop edge;
    those files still receive the white-canvas cleanup above.
    """
    result = image.convert("RGBA")
    width, height = result.size
    if width < 3 or height < 3:
        return result, False

    pixels = result.load()
    step = max(1, max(width, height) // 240)
    edges: dict[str, list[tuple[int, int, int]]] = {"top": [], "bottom": [], "left": [], "right": []}

    def add_sample(edge: str, x: int, y: int) -> None:
        red, green, blue, alpha = pixels[x, y]
        if alpha >= 245:
            edges[edge].append((red, green, blue))

    for x in range(0, width, step):
        add_sample("top", x, 0)
        add_sample("bottom", x, height - 1)
    for y in range(0, height, step):
        add_sample("left", 0, y)
        add_sample("right", width - 1, y)

    samples = [color for values in edges.values() for color in values]
    if len(samples) < 16:
        return result, False

    # Quantizing the edge colors makes a lightly shaded JPG background behave
    # like one candidate while preventing a detailed photo edge from becoming
    # a false background candidate.
    # Screenshots and compressed JPGs can vary by several shades across the
    # same gray or colored canvas, so use a broad enough bucket to keep that
    # one outer background together before the conservative flood-fill step.
    bucket_size = 32
    buckets: dict[tuple[int, int, int], list[tuple[int, int, int]]] = {}
    for color in samples:
        bucket = tuple(channel // bucket_size for channel in color)
        buckets.setdefault(bucket, []).append(color)
    dominant = max(buckets.values(), key=len)
    if len(dominant) / len(samples) < minimum_edge_coverage:
        return result, False
    candidate = tuple(round(sum(color[index] for color in dominant) / len(dominant)) for index in range(3))

    # Let a mildly varied colored canvas remain one background, but keep the
    # tolerance conservative enough that adjacent logo artwork is not eaten.
    seed_tolerance = 44.0
    supporting = [color for color in samples if _rgb_distance(color, candidate) <= seed_tolerance]
    if len(supporting) / len(samples) < minimum_edge_coverage:
        return result, False
    edge_matches = {
        edge: sum(_rgb_distance(color, candidate) <= seed_tolerance for color in values)
        for edge, values in edges.items()
    }
    edges_with_background = sum(
        bool(edges[edge]) and edge_matches[edge] / len(edges[edge]) >= 0.36
        for edge in edges
    )
    if edges_with_background < 3:
        return result, False

    distances = sorted(_rgb_distance(color, candidate) for color in supporting)
    percentile_index = min(len(distances) - 1, round((len(distances) - 1) * 0.90))
    # A photographed or screen-captured canvas often shades from light gray
    # at one edge to darker gray at another.  Permit that gradual variation,
    # but retain near-white artwork (for example, a White Ink version) as a
    # foreground mark rather than swallowing it into a gray background.
    neutral_canvas = max(candidate) - min(candidate) <= 42
    maximum_tolerance = 158.0 if neutral_canvas else 108.0
    minimum_tolerance = 122.0 if neutral_canvas else 102.0
    tolerance = min(maximum_tolerance, max(minimum_tolerance, distances[percentile_index] + 28.0))

    seen = bytearray(width * height)
    pending: deque[tuple[int, int]] = deque()

    def is_background(x: int, y: int) -> bool:
        red, green, blue, alpha = pixels[x, y]
        channels = (red, green, blue)
        protected_white_artwork = min(channels) >= 232 and max(channels) - min(channels) <= 28
        return (
            alpha > 0
            and not protected_white_artwork
            and _rgb_distance(channels, candidate) <= tolerance
        )

    def add_if_background(x: int, y: int) -> None:
        index = y * width + x
        if seen[index] or not is_background(x, y):
            return
        seen[index] = 1
        pending.append((x, y))

    for x in range(width):
        add_if_background(x, 0)
        add_if_background(x, height - 1)
    for y in range(1, height - 1):
        add_if_background(0, y)
        add_if_background(width - 1, y)

    removed = 0
    while pending:
        x, y = pending.popleft()
        red, green, blue, _alpha = pixels[x, y]
        pixels[x, y] = (red, green, blue, 0)
        removed += 1
        if x:
            add_if_background(x - 1, y)
        if x + 1 < width:
            add_if_background(x + 1, y)
        if y:
            add_if_background(x, y - 1)
        if y + 1 < height:
            add_if_background(x, y + 1)

    alpha_box = result.getchannel("A").getbbox()
    if alpha_box and alpha_box != (0, 0, width, height):
        result = result.crop(alpha_box)
        removed += 1
    return result, bool(removed)


def prepare_artwork_for_preview(
    source: str | PathLike[str] | PILImage.Image,
) -> tuple[PILImage.Image, bool]:
    """Return a clean transparent artwork image for an in-app preview.

    This uses the exact same conservative edge cleanup as the outsourced
    report cover.  Keeping it in one place means a saved job logo cannot look
    clean in the PDF while still showing a rectangular canvas on Current Event.
    The source upload is only read; it is never changed here.
    """
    if isinstance(source, PILImage.Image):
        image = source.copy()
    else:
        with PILImage.open(source) as opened:
            image = ImageOps.exif_transpose(opened).convert("RGBA")

    # An existing alpha channel is an intentional artwork boundary. Trim its
    # empty outer padding, but do not reinterpret a colored design at the edge
    # as a removable background a second time.
    already_transparent = image.getchannel("A").getextrema()[0] < 255
    image, trimmed_white = remove_outer_near_white_background(image)
    if already_transparent:
        removed_color = False
    else:
        image, removed_color = remove_outer_edge_background(image)
    return image, bool(trimmed_white or removed_color)


def prepare_artwork_for_report(
    source: str | PathLike[str] | PILImage.Image,
    destination: str | PathLike[str],
) -> bool:
    """Save a clean transparent PNG for the report without changing the upload.

    The saved event copy is intentionally always PNG, even when the customer
    supplied a JPG or WEBP, because transparency is required for a clean PDF
    cover.  ``True`` means background/padding cleanup changed the visible
    artwork.
    """
    image, cleaned = prepare_artwork_for_preview(source)
    target = destination if isinstance(destination, str) else str(destination)
    PILImage.Image.save(image, target, format="PNG", optimize=True)
    return cleaned
