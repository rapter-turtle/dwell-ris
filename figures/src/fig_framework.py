"""Fig. 1 전체 구조도. 실행하면 figures/fig_framework.pdf 와 미리보기 png 를 만든다."""
import os

import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse, FancyArrowPatch, FancyBboxPatch

import figstyle as S

S.apply()
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..')

W, H = S.TEXT_W, 2.85
fig = plt.figure(figsize=(W, H))
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(0, H)
ax.axis('off')

YB, YT = 0.42, 2.52  # 패널 아래, 위


def panel(x, w, title, ec, fc):
    ax.add_patch(FancyBboxPatch((x, YB), w, YT - YB, boxstyle='round,pad=0,rounding_size=0.06',
                                fc=fc, ec=ec, lw=0.9))
    ax.text(x + w / 2, YT - 0.12, title, ha='center', va='center', fontsize=7.6,
            fontweight='bold', color=ec)


def box(x, y0, w, y1, text, ec, fc='white', fs=7.0, style='normal', color='black', lw=0.7):
    ax.add_patch(FancyBboxPatch((x, y0), w, y1 - y0, boxstyle='round,pad=0,rounding_size=0.04',
                                fc=fc, ec=ec, lw=lw))
    ax.text(x + w / 2, (y0 + y1) / 2, text, ha='center', va='center', fontsize=fs,
            linespacing=1.25, style=style, color=color)


def arrow(p0, p1, color='black', lw=0.9, style='-|>', ls='-', rad=0.0):
    ax.add_patch(FancyArrowPatch(p0, p1, arrowstyle=style, mutation_scale=7, lw=lw,
                                 color=color, linestyle=ls,
                                 connectionstyle=f'arc3,rad={rad}', shrinkA=0, shrinkB=0))


# 패널 위치
P1x, P1w = 0.05, 1.62
P2x, P2w = 1.93, 2.20
P3x, P3w = 4.39, 1.17
P4x, P4w = 5.82, 1.29

# (1) 선박과 작동기
panel(P1x, P1w, 'Vessel with gear shift dwell', S.C_PLANT, S.C_PLANT_BG)
bx, bw = P1x + 0.09, P1w - 0.18
box(bx, 1.86, bw, 2.26, 'Gear shift $g\\in\\{\\mathrm{R},\\mathrm{N},\\mathrm{F}\\}$\nheld $D_g=6$ steps (3 s)', S.C_PLANT)
box(bx, 1.36, bw, 1.76, 'Sterndrive steering $\\delta$\nrate limited, no dwell', S.C_PLANT)
box(bx, 0.86, bw, 1.26, 'Disturbance $d\\in\\mathcal{W}$\nwind and current', S.C_PLANT)
ax.text(P1x + P1w / 2, 0.64, 'Locked in N, thrust and\nsteering force both vanish',
        ha='center', va='center', fontsize=6.8, style='italic', color='#9B2222', linespacing=1.2)

# (2) Stage 1
panel(P2x, P2w, 'Stage 1 offline, dwell indexed RIS', S.C_OFF, S.C_OFF_BG)
bx, bw = P2x + 0.09, P2w - 0.18
box(bx, 1.90, bw, 2.26, 'Augmented state\n$\\xi=(\\tilde{x},\\tilde{y},\\tilde{\\psi},u,v,r,\\delta,g,\\tau_g)$', S.C_OFF)
half = (bw - 0.08) / 2
box(bx, 1.42, half, 1.80, 'Protagonist $\\pi_h$\nmax over $C(\\xi)$', S.C_OFF)
box(bx + half + 0.08, 1.42, half, 1.80, 'Adversary $\\mu$\nmin over $\\mathcal{W}$', S.C_OFF)
box(bx, 0.94, bw, 1.32, 'Twin safety critics $Q_{h,1}$, $Q_{h,2}$\none step discounted backup, masked', S.C_OFF)
# 안전집합 족 그림 (경과 dwell 이 길수록 작아짐)
cx, cy = bx + 0.40, 0.64
for (a, b, al) in [(0.74, 0.34, 0.16), (0.50, 0.23, 0.30), (0.26, 0.12, 0.55)]:
    ax.add_patch(Ellipse((cx, cy), a, b, fc=S.C_OFF, alpha=al, ec=S.C_OFF, lw=0.6))
# 각 고리 안쪽에 경과 dwell 값 표시
for (dx, lab, col) in [(0.31, '0', S.C_OFF), (0.19, '3', S.C_OFF), (0.0, '5', 'white')]:
    ax.text(cx + dx, cy, lab, fontsize=6.2, color=col, va='center', ha='center')
ax.text(cx, cy + 0.235, '$\\tau_g=0,\\ 3,\\ 5$', fontsize=6.3, color=S.C_OFF, va='center', ha='center')
ax.text(bx + bw - 0.02, 0.67, 'Safe set family\n$\\mathcal{S}(g,\\tau_g)$, smaller\ndeeper in a lock',
        ha='right', va='center', fontsize=6.8, linespacing=1.2)

# (3) Stage 2
panel(P3x, P3w, 'Stage 2 offline', S.C_OFF, S.C_OFF_BG)
bx, bw = P3x + 0.08, P3w - 0.16
box(bx, 1.70, bw, 2.24, 'Task policy $\\pi_t$\nsoft actor critic', S.C_OFF)
box(bx, 1.06, bw, 1.58, 'Lagrangian on\n$q\\geq\\tau$, $\\lambda\\geq\\lambda_{\\min}$', S.C_OFF)
ax.text(P3x + P3w / 2, 0.74, 'safety critic $q$\nfrozen from\nStage 1', ha='center', va='center',
        fontsize=6.6, color=S.C_GRAY, linespacing=1.2)

# (4) 온보드
panel(P4x, P4w, 'Onboard, every 0.5 s', S.C_ON, S.C_ON_BG)
bx, bw = P4x + 0.08, P4w - 0.16
box(bx, 1.88, bw, 2.26, 'EKF state $\\xi_k$\nat 10 Hz', S.C_ON)
box(bx, 1.16, bw, 1.76, 'Shield\napply $a_t$ if $q(\\xi,a_t)\\geq\\tau$\nelse apply $\\pi_h(\\xi)$', S.C_ON)
box(bx, 0.52, bw, 1.04, 'PLC and\nservo cylinders', S.C_ON)
arrow((P4x + P4w / 2, 1.88), (P4x + P4w / 2, 1.76), color=S.C_ON)
arrow((P4x + P4w / 2, 1.16), (P4x + P4w / 2, 1.04), color=S.C_ON)

# 패널 사이 화살표
ym = 1.56
arrow((P1x + P1w, ym), (P2x, ym))
ax.text((P1x + P1w + P2x) / 2, ym + 0.10, 'model', ha='center', fontsize=6.5)
arrow((P2x + P2w, ym), (P3x, ym))
ax.text((P2x + P2w + P3x) / 2, ym + 0.10, '$q$, $\\pi_h$', ha='center', fontsize=6.5)
arrow((P3x + P3w, ym), (P4x, ym))
ax.text((P3x + P3w + P4x) / 2, ym + 0.10, '$\\pi_t$', ha='center', fontsize=6.5)

# Stage 1 에서 온보드 shield 로 가는 경로 (위쪽)
yt = 2.68
xs = P2x + P2w * 0.75
xe = P4x + P4w * 0.30
ax.plot([xs, xs, xe], [YT, yt, yt], color='black', lw=0.9)
arrow((xe, yt), (xe, YT), color='black')
ax.text((xs + xe) / 2, yt + 0.07, 'certified $q$, $\\pi_h$, and margin $\\tau$', ha='center', fontsize=6.5)

# 아래쪽 폐루프 (명령, 측정)
y1, y2 = 0.27, 0.11
x_on = P4x + P4w * 0.5
x_pl = P1x + P1w * 0.5
ax.plot([x_on, x_on, x_pl + 0.15], [YB, y1, y1], color=S.C_ON, lw=0.9)
arrow((x_pl + 0.15, y1), (x_pl + 0.15, YB), color=S.C_ON)
ax.text((x_on + x_pl) / 2, y1 + 0.05, 'gear and steering commands to the vessel', ha='center',
        fontsize=6.5, color=S.C_ON)
ax.plot([x_pl - 0.15, x_pl - 0.15, x_on + 0.25], [YB, y2, y2], color=S.C_PLANT, lw=0.9)
arrow((x_on + 0.25, y2), (x_on + 0.25, YB), color=S.C_PLANT)
ax.text((x_on + x_pl) / 2, y2 - 0.075, 'GNSS and IMU measurements', ha='center', fontsize=6.5,
        color=S.C_PLANT, va='center')

fig.savefig(os.path.join(OUT, 'fig_framework.pdf'))
fig.savefig(os.path.join(HERE, 'preview_fig_framework.png'), dpi=220)
print('saved')
