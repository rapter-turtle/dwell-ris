"""Fig. 2(a): 추력 deadzone 과 기어 시프트 dwell 설명 그림.

왼쪽: 요청 추력에 대해 실제로 낼 수 있는 추력 (아이들 최소 추력 때문에 0 주변이 deadzone).
오른쪽: 시간에 따른 기어 (R, N, F) 와 시프트 뒤의 잠금 구간 (dwell), 잠금 중에 무시되는 요청.
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import figstyle as fs  # noqa: E402

fs.apply()
OUT = os.path.join(HERE, '..')

C_LOCK = '#F3D9BF'
C_DZ = '#E3E3E3'


def make():
    fig = plt.figure(figsize=(3.45, 1.0))
    ax1 = fig.add_axes([0.115, 0.22, 0.355, 0.72])
    ax2 = fig.add_axes([0.605, 0.22, 0.375, 0.72])

    # ---- deadzone: requested thrust -> available thrust ----
    fmin, fmax = 0.35, 1.0
    ax1.axvspan(-fmin, fmin, color=C_DZ, lw=0)
    x = np.linspace(-fmax, fmax, 401)
    y = np.where(np.abs(x) < fmin, 0.0, x)
    for seg in (x < -fmin, np.abs(x) < fmin, x > fmin):
        ax1.plot(x[seg], y[seg], color=fs.C_PLANT, lw=1.2)
    ax1.plot([-fmin, fmin], [-fmin, fmin], color=fs.C_PLANT, lw=0.6, ls=':')
    ax1.plot([fmin, -fmin], [fmin, -fmin], 'o', ms=2.6, color=fs.C_PLANT, mfc=fs.C_PLANT)
    ax1.plot([fmin, -fmin], [0, 0], 'o', ms=2.6, color=fs.C_PLANT, mfc='white')
    ax1.text(0, 0.62, 'deadzone', ha='center', va='center', fontsize=6.5, color=fs.C_GRAY)
    ax1.text(fmin + 0.05, fmin - 0.02, 'idle', ha='left', va='top', fontsize=6.2, color=fs.C_PLANT)
    ax1.set_xlim(-fmax, fmax)
    ax1.set_ylim(-fmax, fmax)
    ax1.set_xticks([])
    ax1.set_yticks([-fmin, 0, fmin])
    ax1.set_yticklabels([r'$-F_{\mathrm{idle}}$', '0', r'$F_{\mathrm{idle}}$'], fontsize=6.2)
    ax1.set_xlabel('requested thrust', fontsize=6.6, labelpad=1)
    ax1.set_ylabel('thrust', fontsize=6.6, labelpad=1)
    ax1.tick_params(length=2, pad=1)

    # ---- dwell: gear over time with lock windows ----
    D = 6
    shifts = [(2, 1), (10, 0), (19, 1)]          # (step, new level) 0=R, 1=N, 2=F
    g0 = 2
    T = 26
    t = np.arange(T + 1)
    g = np.full(T + 1, g0, dtype=float)
    for k, lv in shifts:
        g[k:] = lv
    for k, _ in shifts:
        ax2.axvspan(k, k + D, color=C_LOCK, lw=0)
    ax2.step(t, g, where='post', color=fs.C_PLANT, lw=1.2)
    # 잠금 중 무시된 요청
    ax2.plot([4.5], [0], marker='x', ms=4, mew=1.0, color='#9B2222')
    ax2.text(5.2, 0.0, 'ignored', fontsize=6.0, color='#9B2222', ha='left', va='center')
    # dwell 길이 표시
    ax2.annotate('', xy=(10, 2.45), xytext=(16, 2.45),
                 arrowprops=dict(arrowstyle='<->', lw=0.6, color=fs.C_OFF, shrinkA=0, shrinkB=0))
    ax2.text(13, 2.55, r'$D_g$', ha='center', va='bottom', fontsize=6.6, color=fs.C_OFF)
    ax2.set_xlim(0, T)
    ax2.set_ylim(-0.4, 2.9)
    ax2.set_yticks([0, 1, 2])
    ax2.set_yticklabels(['R', 'N', 'F'], fontsize=6.4)
    ax2.set_xticks([])
    ax2.set_xlabel('time step', fontsize=6.6, labelpad=1)
    ax2.set_ylabel('gear', fontsize=6.6, labelpad=1)
    ax2.tick_params(length=2, pad=1)
    for a in (ax1, ax2):
        for s in ('top', 'right'):
            a.spines[s].set_visible(False)

    pdf = os.path.join(OUT, 'fig_gear.pdf')
    fig.savefig(pdf)
    fig.savefig(os.path.join(HERE, 'render_fig_gear.png'), dpi=300)
    return pdf


if __name__ == '__main__':
    print(make())
