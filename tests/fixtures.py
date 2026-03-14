"""Test fixture generators — create minimal test files without external dependencies."""
import os
import struct
import tempfile

TEMP_DIR = os.environ.get("TEMP", tempfile.gettempdir())


def create_test_bmp(path=None, width=4, height=4, color=(255, 0, 0)):
    """Create a minimal 24-bit BMP image file. Returns the file path."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture.bmp")

    row_size = (width * 3 + 3) & ~3  # rows padded to 4-byte boundary
    pixel_data_size = row_size * height
    file_size = 54 + pixel_data_size

    r, g, b = color
    with open(path, "wb") as f:
        # BMP header (14 bytes)
        f.write(b"BM")
        f.write(struct.pack("<I", file_size))
        f.write(struct.pack("<HH", 0, 0))
        f.write(struct.pack("<I", 54))
        # DIB header (40 bytes)
        f.write(struct.pack("<I", 40))
        f.write(struct.pack("<i", width))
        f.write(struct.pack("<i", height))
        f.write(struct.pack("<HH", 1, 24))
        f.write(struct.pack("<I", 0))
        f.write(struct.pack("<I", pixel_data_size))
        f.write(struct.pack("<i", 2835))
        f.write(struct.pack("<i", 2835))
        f.write(struct.pack("<II", 0, 0))
        # Pixel data (BGR order, bottom-up)
        row = bytes([b, g, r]) * width
        row += b"\x00" * (row_size - width * 3)
        for _ in range(height):
            f.write(row)
    return path


def create_test_bmp_2(path=None, width=4, height=4):
    """Create a second BMP with a different color (blue)."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture_2.bmp")
    return create_test_bmp(path, width, height, color=(0, 0, 255))


def create_test_wav(path=None, duration_ms=100, sample_rate=8000):
    """Create a minimal WAV audio file. Returns the file path."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture.wav")

    num_samples = int(sample_rate * duration_ms / 1000)
    data_size = num_samples * 2  # 16-bit mono
    file_size = 36 + data_size

    with open(path, "wb") as f:
        # RIFF header
        f.write(b"RIFF")
        f.write(struct.pack("<I", file_size))
        f.write(b"WAVE")
        # fmt chunk
        f.write(b"fmt ")
        f.write(struct.pack("<I", 16))
        f.write(struct.pack("<HHIIHH", 1, 1, sample_rate, sample_rate * 2, 2, 16))
        # data chunk — silence
        f.write(b"data")
        f.write(struct.pack("<I", data_size))
        f.write(b"\x00" * data_size)
    return path


def create_test_avi(path=None):
    """Create a minimal AVI video file (1 frame, 2x2 pixels). Returns the file path."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture.avi")

    # Minimal uncompressed AVI: 2x2 pixels, 1 frame, 24-bit BGR
    width, height = 2, 2
    row_size = (width * 3 + 3) & ~3
    frame_size = row_size * height
    # Pad frame to even size
    padded_frame = frame_size if frame_size % 2 == 0 else frame_size + 1

    def chunk(tag, data):
        d = data if isinstance(data, bytes) else data
        pad = b"\x00" if len(d) % 2 else b""
        return tag + struct.pack("<I", len(d)) + d + pad

    def list_chunk(tag, sub_tag, data):
        inner = sub_tag + data
        return tag + struct.pack("<I", len(inner)) + inner

    # Frame data (blue pixels)
    row = b"\xff\x00\x00" * width + b"\x00" * (row_size - width * 3)
    frame_data = row * height

    # Build AVI
    avih = struct.pack("<IIIIIIIIIIIIII",
        66667,      # dwMicroSecPerFrame (15fps)
        0, 0, 0,    # dwMaxBytesPerSec, dwPaddingGranularity, dwFlags
        1,          # dwTotalFrames
        0,          # dwInitialFrames
        1,          # dwStreams
        frame_size, # dwSuggestedBufferSize
        width, height, 0, 0, 0, 0)

    strh = struct.pack("<4s4sIHHIIIIIIIIHHHH",
        b"vids", b"DIB ", 0, 0, 0, 0, 1, 15, 0, 1, frame_size, 0, 0, 0, 0, 0, 0)

    bmi = struct.pack("<IiiHHIIiiII",
        40, width, height, 1, 24, 0, frame_size, 0, 0, 0, 0)

    strl_data = chunk(b"strh", strh) + chunk(b"strf", bmi)
    strl = list_chunk(b"RIFF", b"strl", strl_data)[4:]  # strip outer RIFF, use LIST
    strl = b"LIST" + struct.pack("<I", len(b"strl" + strl_data)) + b"strl" + strl_data

    hdrl_data = chunk(b"avih", avih) + strl
    hdrl = b"LIST" + struct.pack("<I", len(b"hdrl" + hdrl_data)) + b"hdrl" + hdrl_data

    movi_data = chunk(b"00dc", frame_data)
    movi = b"LIST" + struct.pack("<I", len(b"movi" + movi_data)) + b"movi" + movi_data

    avi_data = hdrl + movi
    with open(path, "wb") as f:
        f.write(b"RIFF")
        f.write(struct.pack("<I", len(b"AVI " + avi_data)))
        f.write(b"AVI ")
        f.write(avi_data)
    return path


def create_test_txt(path=None, content="Test OLE content."):
    """Create a small text file for OLE embedding. Returns the file path."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture.txt")
    with open(path, "w") as f:
        f.write(content)
    return path


def create_test_pptx(path=None):
    """Save the active PowerPoint presentation to a temp file. Returns the path.
    Caller must have an active presentation open."""
    if path is None:
        path = os.path.join(TEMP_DIR, "test_fixture.pptx")
    return path


def cleanup_files(*paths):
    """Remove temp files, ignoring errors."""
    for p in paths:
        try:
            if p and os.path.isfile(p):
                os.remove(p)
        except OSError:
            pass
