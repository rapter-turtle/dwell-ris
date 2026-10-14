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
sd_crop = sd.crop((int(130 * k), int(470 * k), int(960 * k), int(1050 * k)))
save(sd_crop, 'sterndrive.jpg', 1400)


def lift_shadows(im, clip=3.0, gamma=0.75):
    """그늘에 묻힌 어두운 부분을 밝힌다. 밝기 채널에 CLAHE 와 감마를 적용하고 색은 유지."""
    import cv2
    import numpy as np
    bgr = cv2.cvtColor(np.asarray(im), cv2.COLOR_RGB2BGR)
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    l = cv2.createCLAHE(clipLimit=clip, tileGridSize=(8, 8)).apply(l)
    lut = np.array([((i / 255.0) ** gamma) * 255 for i in range(256)], dtype=np.uint8)
    l = cv2.LUT(l, lut)
    out = cv2.cvtColor(cv2.merge((l, a, b)), cv2.COLOR_LAB2BGR)
    return Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))


save(lift_shadows(sd_crop), 'sterndrive_lifted.jpg', 1400)
