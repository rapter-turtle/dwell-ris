"""Fig. 2 선박과 하드웨어. 실행하면 figures/fig_vessel.pdf 와 미리보기 png 를 만든다.

(a) 선박 사진, (b) 선미 기관실 사진 (시프트, 조향 실린더 라벨), (c) 하드웨어 연결 구성도.
사진은 photos/ 폴더에 둔다. (b) 라벨 위치는 'Ship mechanical system.pptx' 7번 슬라이드의
원래 화살표 끝점을 사진 좌표로 되돌려 계산한 값이다.
"""
import os

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from PIL import Image

import figstyle as S

S.apply()
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..')
PH = os.path.join(HERE, 'photos')

W, H = S.TEXT_W, 2.45
fig = plt.figure(figsize=(W, H))

# 패널 배치 [left, bottom, width, height] (그림 비율 단위)
def axes_in(x, y, w, h):
    return fig.add_axes([x / W, y / H, w / W, h / H])


# (a) 선박 사진
img = Image.open(os.path.join(PH, 'vessel_side.jpg'))
img_a = img.crop((140, 300, 1340, 1040))  # 배 위주로 자름 (1200 x 740)
wa = 2.62
ha = wa * img_a.height / img_a.width
axa = axes_in(0.02, 0.26 + (2.05 - ha) / 2, wa, ha)
axa.imshow(img_a)
axa.axis('off')

# (b) 선미 기관실 사진
img = Image.open(os.path.join(PH, 'stern_bay.jpg'))
x0, y0, x1, y1 = 300, 360, 760, 960
img_b = img.crop((x0, y0, x1, y1))
hb = 2.05
wb = hb * img_b.width / img_b.height
axb = axes_in(0.02 + wa + 0.10, 0.26, wb, hb)
axb.imshow(img_b)
axb.axis('off')

# 슬라이드 화살표 끝점 (원본 사진 픽셀 좌표)
targets = {
    'shift': (638, 691),
    'steer': (413, 539),
    'pivot': (485, 858),
}


def lab(key, text, txy):
    tx, ty = targets[key]
    axb.annotate(text, xy=(tx - x0, ty - y0), xytext=txy, textcoords='data',
                 fontsize=6.4, ha='center', va='center', linespacing=1.15,
                 bbox=dict(boxstyle='round,pad=0.18', fc='white', ec=S.C_PLANT, lw=0.6, alpha=0.92),
                 arrowprops=dict(arrowstyle='-|>', color='#FFD000', lw=1.0, mutation_scale=6,
                                 shrinkA=0, shrinkB=1))


lab('steer', 'Steering servo\ncylinder', (120, 40))
lab('shift', 'Shift servo\ncylinder (R/N/F)', (345, 150))
lab('pivot', 'Drive steering\npivot', (150, 560))

# (c) 하드웨어 연결 구성도
xc = 0.02 + wa + 0.10 + wb + 0.12
wc = W - xc - 0.02
axc = axes_in(xc, 0.26, wc, 2.05)
axc.set_xlim(0, wc)
axc.set_ylim(0, 2.05)
axc.axis('off')


def hbox(x, y, w, h, text, ec, fc='white', ls='-', color='black', fs=6.6):
    axc.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0,rounding_size=0.04',
                                 fc=fc, ec=ec, lw=0.7, linestyle=ls))
    axc.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=fs,
             linespacing=1.2, color=color)


def link(p0, p1, text=None, toff=(0.0, 0.0), color='black', ls='-'):
    axc.add_patch(FancyArrowPatch(p0, p1, arrowstyle='-|>', mutation_scale=6, lw=0.8,
                                  color=color, linestyle=ls, shrinkA=0, shrinkB=0))
    if text:
        mx, my = (p0[0] + p1[0]) / 2 + toff[0], (p0[1] + p1[1]) / 2 + toff[1]
        axc.text(mx, my, text, fontsize=5.9, ha='center', va='center', color=S.C_GRAY)


bw_ = (wc - 0.12) / 2
# 센서
hbox(0.0, 1.62, bw_, 0.38, 'Here3 GNSS\nreceiver', S.C_PLANT, S.C_PLANT_BG)
hbox(bw_ + 0.12, 1.62, bw_, 0.38, 'Pixhawk Cube\nwith IMU', S.C_PLANT, S.C_PLANT_BG)
link((bw_, 1.81), (bw_ + 0.12, 1.81))
# 온보드 컴퓨터
hbox(0.0, 1.00, wc, 0.42, 'Onboard computer, ROS 2\nEKF at 10 Hz, shielded controller every 0.5 s', S.C_ON, S.C_ON_BG)
link((bw_ + 0.12 + bw_ / 2, 1.62), (bw_ + 0.12 + bw_ / 2, 1.42), 'serial', (0.17, 0.0))
# PLC (폭 전체로 두어 작동기로 곧게 내려가게 함)
hbox(0.0, 0.58, wc, 0.26, 'LS XGB PLC', S.C_GRAY, S.C_GRAY_BG)
link((wc / 2, 1.00), (wc / 2, 0.84), 'Ethernet', (0.25, 0.0))
# 작동기
aw = (wc - 0.16) / 3
labels = ['Shift servo\ncylinder\nR/N/F', 'Steering servo\ncylinder\n$\\delta$', 'Engine\nspeed']
for i, t in enumerate(labels):
    hbox(i * (aw + 0.08), 0.0, aw, 0.40, t, S.C_PLANT, S.C_PLANT_BG, fs=6.2)
for i, t in enumerate(['EtherCAT', 'EtherCAT', 'CAN']):
    cx = i * (aw + 0.08) + aw / 2
    link((cx, 0.58), (cx, 0.40))
    axc.text(cx + 0.04, 0.49, t, fontsize=5.9, ha='left', va='center', color=S.C_GRAY)

fig.text((0.02 + wa / 2) / W, 0.07, '(a)', ha='center', fontsize=7.5)
fig.text((0.02 + wa + 0.10 + wb / 2) / W, 0.07, '(b)', ha='center', fontsize=7.5)
fig.text((xc + wc / 2) / W, 0.07, '(c)', ha='center', fontsize=7.5)

fig.savefig(os.path.join(OUT, 'fig_vessel.pdf'))
fig.savefig(os.path.join(HERE, 'preview_fig_vessel.png'), dpi=220)
print('saved', round(wa, 2), round(wb, 2), round(wc, 2))
