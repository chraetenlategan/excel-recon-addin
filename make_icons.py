"""Generate the add-in icons (green tile + white check) at 16/32/64/80 px.
No third-party deps — writes PNGs by hand via zlib. Run once; commit the PNGs.
"""
import zlib, struct, os

GREEN = (33, 115, 70)   # Excel green
WHITE = (255, 255, 255)

def check_mask(size):
    """Return a set of (x,y) pixels forming a bold checkmark, scaled to size."""
    pts = set()
    s = size
    thick = max(1, size // 10)
    # Two strokes of a check: short down-right, then long up-right.
    def stroke(x0, y0, x1, y1):
        steps = int(max(abs(x1 - x0), abs(y1 - y0)) * 4) + 1
        for i in range(steps + 1):
            t = i / steps
            x = x0 + (x1 - x0) * t
            y = y0 + (y1 - y0) * t
            for dx in range(-thick, thick + 1):
                for dy in range(-thick, thick + 1):
                    px, py = round(x + dx), round(y + dy)
                    if 0 <= px < s and 0 <= py < s:
                        pts.add((px, py))
    stroke(s * 0.24, s * 0.52, s * 0.42, s * 0.70)
    stroke(s * 0.42, s * 0.70, s * 0.76, s * 0.30)
    return pts

def make_png(size, path):
    check = check_mask(size)
    raw = bytearray()
    for y in range(size):
        raw.append(0)  # filter type 0
        for x in range(size):
            r, g, b = WHITE if (x, y) in check else GREEN
            raw += bytes((r, g, b))
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)  # 8-bit RGB
    idat = zlib.compress(bytes(raw), 9)
    png = sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)
    print("wrote", path, size)

here = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(here, exist_ok=True)
for s in (16, 32, 64, 80):
    make_png(s, os.path.join(here, f"icon-{s}.png"))
