"""Prepare installer assets using only the Python standard library."""
from hashlib import sha256
from pathlib import Path
import struct
from urllib.request import urlopen

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / 'build/windows'
MODEL_URL = 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/4.1.0/rus.traineddata'
MODEL_SHA256 = 'e16e5e036cce1d9ec2b00063cf8b54472625b9e14d893a169e2b0dedeb4df225'


def build():
    BUILD.mkdir(parents=True, exist_ok=True)
    model = BUILD / 'rus.traineddata'
    if not model.exists() or sha256(model.read_bytes()).hexdigest() != MODEL_SHA256:
        with urlopen(MODEL_URL, timeout=90) as response:
            data = response.read()
        if sha256(data).hexdigest() != MODEL_SHA256:
            raise ValueError('Russian OCR model checksum does not match')
        model.write_bytes(data)

    # ICO permits embedded PNG images. Preserve the original owl artwork.
    png = (ROOT / 'icons/owl-lawyer-256.png').read_bytes()
    width, height = struct.unpack('>II', png[16:24])
    assert (width, height) == (256, 256)
    header = struct.pack('<HHH', 0, 1, 1)
    entry = struct.pack('<BBBBHHII', 0, 0, 0, 0, 1, 32, len(png), 22)
    (BUILD / 'owl-lawyer.ico').write_bytes(header + entry + png)
    print(f'Installer assets ready: {BUILD}')


if __name__ == '__main__':
    build()
