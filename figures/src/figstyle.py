"""논문 그림 공통 스타일 (IEEE 2단, Times 계열 글꼴)."""
import matplotlib as mpl

COL_W = 3.5    # 단일 단 폭 [in]
TEXT_W = 7.16  # 2단 전체 폭 [in]

# 색상 (흑백 인쇄에서도 구분되도록 명도 차이를 둠)
C_PLANT = '#2E5A87'      # 선박, 하드웨어
C_PLANT_BG = '#E8EFF6'
C_OFF = '#A65E1A'        # 오프라인 학습
C_OFF_BG = '#FBEFE3'
C_ON = '#2F7D4F'         # 온보드
C_ON_BG = '#E7F3EC'
C_GRAY = '#6B6B6B'
C_GRAY_BG = '#F2F2F2'


def apply():
    mpl.rcParams.update({
        'font.family': 'serif',
        'font.serif': ['STIXGeneral', 'Times New Roman', 'DejaVu Serif'],
        'mathtext.fontset': 'stix',
        'font.size': 7.5,
        'axes.linewidth': 0.6,
        'pdf.fonttype': 42,
        'ps.fonttype': 42,
        'savefig.dpi': 300,
    })
