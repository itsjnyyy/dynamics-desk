import struct, zlib, os

ACCENT_A = (108, 192, 245)   # #6cc0f5
ACCENT_B = (74, 157, 219)    # #4a9ddb

def lerp(a, b, t):
    return a + (b - a) * t

def render(size):
    R = size * (11 / 38)          # corner radius, matches app logo proportions
    s = size / 38
    plus_half = (2.5 * s) / 2
    h_x0, h_x1, h_y = 12 * s, 26 * s, 19 * s
    v_y0, v_y1, v_x = 12 * s, 26 * s, 19 * s

    pixels = bytearray()
    for y in range(size):
        row = bytearray()
        for x in range(size):
            cx, cy = x + 0.5, y + 0.5

            # rounded-rect mask
            inside = True
            if cx < R and cy < R:
                inside = (cx - R) ** 2 + (cy - R) ** 2 <= R * R
            elif cx > size - R and cy < R:
                inside = (cx - (size - R)) ** 2 + (cy - R) ** 2 <= R * R
            elif cx < R and cy > size - R:
                inside = (cx - R) ** 2 + (cy - (size - R)) ** 2 <= R * R
            elif cx > size - R and cy > size - R:
                inside = (cx - (size - R)) ** 2 + (cy - (size - R)) ** 2 <= R * R

            if not inside:
                row += bytes((0, 0, 0, 0))
                continue

            t = (cx + cy) / (2 * size)
            r = int(lerp(ACCENT_A[0], ACCENT_B[0], t))
            g = int(lerp(ACCENT_A[1], ACCENT_B[1], t))
            b = int(lerp(ACCENT_A[2], ACCENT_B[2], t))

            on_h = (h_x0 - plus_half <= cx <= h_x1 + plus_half) and (h_y - plus_half <= cy <= h_y + plus_half)
            on_v = (v_x - plus_half <= cx <= v_x + plus_half) and (v_y0 - plus_half <= cy <= v_y1 + plus_half)

            if on_h or on_v:
                row += bytes((255, 255, 255, 255))
            else:
                row += bytes((r, g, b, 255))
        pixels += bytes((0,)) + row  # filter type 0 per scanline
    return bytes(pixels)

def write_png(path, size):
    raw = render(size)
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xffffffff)
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    idat = zlib.compress(raw, 9)
    with open(path, "wb") as f:
        f.write(sig)
        f.write(chunk(b"IHDR", ihdr))
        f.write(chunk(b"IDAT", idat))
        f.write(chunk(b"IEND", b""))
    return raw

def write_ico(path, sizes):
    entries = []
    images = []
    for size in sizes:
        png_path = path + f".{size}.tmp.png"
        write_png(png_path, size)
        with open(png_path, "rb") as f:
            data = f.read()
        os.remove(png_path)
        images.append(data)
        entries.append(size)

    with open(path, "wb") as f:
        f.write(struct.pack("<HHH", 0, 1, len(sizes)))
        offset = 6 + 16 * len(sizes)
        for size, data in zip(entries, images):
            w = size if size < 256 else 0
            h = size if size < 256 else 0
            f.write(struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(data), offset))
            offset += len(data)
        for data in images:
            f.write(data)

if __name__ == "__main__":
    base = os.path.dirname(os.path.abspath(__file__))
    write_png(os.path.join(base, "icon.png"), 256)
    write_ico(os.path.join(base, "icon.ico"), [16, 24, 32, 48, 64, 128, 256])
    print("done")
