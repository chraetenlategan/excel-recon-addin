"""Generate the add-in icons at 16/32/64/80 px.

A rounded Excel-green tile with a white check, drawn on a transparent
background with 4x supersampled anti-aliasing so the small sizes stay crisp
in the ribbon. No third-party deps — writes RGBA PNGs by hand via zlib.
Run once; commit the PNGs.
"""
import zlib, struct, os, math

GREEN = (33, 115, 70)     # Excel green
WHITE = (255, 255, 255)
SS = 4                    # supersampling factor per axis

# Geometry in unit (0..1) coordinates, scaled to whatever size is rendered.
RADIUS = 0.22             # tile corner radius
CHECK = [(0.28, 0.53), (0.44, 0.69), (0.75, 0.34)]
STROKE = 0.075            # half-width of the check stroke


def _rounded_rect_cover(x, y, r):
    """1 inside the unit rounded square, 0 outside (sampled, so no blending)."""
    cx = min(max(x, r), 1 - r)
    cy = min(max(y, r), 1 - r)
    if x >= r and x <= 1 - r:
        return 1.0 if 0 <= y <= 1 else 0.0
    if y >= r and y <= 1 - r:
        return 1.0 if 0 <= x <= 1 else 0.0
    return 1.0 if math.hypot(x - cx, y - cy) <= r else 0.0


def _seg_dist(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
    t = min(max(t, 0.0), 1.0)
    return math.hypot(px - (ax + dx * t), py - (ay + dy * t))


def _check_cover(x, y):
    """1 inside the check (round caps and joins), else 0."""
    for i in range(len(CHECK) - 1):
        ax, ay = CHECK[i]
        bx, by = CHECK[i + 1]
        if _seg_dist(x, y, ax, ay, bx, by) <= STROKE:
            return 1.0
    return 0.0


def _pixel(px, py, size):
    """Average SS x SS samples into one RGBA pixel."""
    tile = check = 0
    for sy in range(SS):
        for sx in range(SS):
            x = (px + (sx + 0.5) / SS) / size
            y = (py + (sy + 0.5) / SS) / size
            t = _rounded_rect_cover(x, y, RADIUS)
            tile += t
            check += t * _check_cover(x, y)
    n = SS * SS
    alpha = tile / n
    if alpha == 0:
        return (0, 0, 0, 0)
    # Colour is green tinted towards white by however much check covered it.
    w = (check / n) / alpha
    rgb = tuple(round(g + (wc - g) * w) for g, wc in zip(GREEN, WHITE))
    return rgb + (round(alpha * 255),)


def make_png(size, path):
    raw = bytearray()
    for y in range(size):
        raw.append(0)  # filter type 0
        for x in range(size):
            raw += bytes(_pixel(x, y, size))

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)  # 8-bit RGBA
    idat = zlib.compress(bytes(raw), 9)
    with open(path, "wb") as f:
        f.write(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b""))
    print("wrote", path, size)


here = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(here, exist_ok=True)
for s in (16, 32, 64, 80, 128):
    make_png(s, os.path.join(here, f"icon-{s}.png"))
