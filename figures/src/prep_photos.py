"""'Ship configuration' 폴더의 사진을 논문 그림용으로 정리한다 (회전 보정, 자르기, 축소)."""
import os

from PIL import Image, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '..', '..', '..', 'Ship configuration')
DST = os.path.join(HERE, 'photos')
os.makedirs(DST, exist_ok=True)


def load(name):
    return ImageOps.exif_transpose(Image.open(os.path.join(SRC, name))).convert('RGB')


def save(im, name, max_w):
    if im.width > max_w:
        im = im.resize((max_w, int(im.height * max_w / im.width)), Image.LANCZOS)
    im.save(os.path.join(DST, name), quality=92)
    print(name, im.size)


# 실험선과 센서 (사용자가 만든 합성 사진)
ap = load('aura_pic2.png')
save(ap, 'vessel_sensors.jpg', 2215)

# Control PC (노트북). 파란 사각형 안쪽만 자름 (축소본 1400 폭 기준 좌표를 원본으로 환산)
s = ap.width / 1400.0
save(ap.crop((int(1100 * s), int(420 * s), int(1252 * s), int(592 * s))), 'control_pc_laptop.jpg', 600)

# 선미 드라이브 (축소본 1050 폭 기준 좌표를 원본으로 환산)
sd = load('KakaoTalk_20261014_105036447_01.jpg')
k = sd.width / 1050.0
save(sd.crop((int(130 * k), int(470 * k), int(960 * k), int(1050 * k))), 'sterndrive.jpg', 1400)
