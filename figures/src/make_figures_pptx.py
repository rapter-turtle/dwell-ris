"""Fig. 1, Fig. 2 를 손으로 고칠 수 있는 pptx 로 만든다 (모든 상자, 글자, 화살표가 개별 도형).

슬라이드 크기 = 논문 그림 크기. PowerPoint 에서 고친 뒤 '내보내기 > PDF' 로 저장하면
figures/ 의 PDF 를 그대로 교체할 수 있다. 좌표는 fig_framework.py, fig_vessel.py 와 같다.
"""
import os
import copy

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(HERE, '..', 'ppt')
PH = os.path.join(HERE, 'photos')
os.makedirs(OUTDIR, exist_ok=True)

FONT = 'Times New Roman'
MATHFONT = 'Cambria Math'

C_PLANT, C_PLANT_BG = '2E5A87', 'E8EFF6'
C_OFF, C_OFF_BG = 'A65E1A', 'FBEFE3'
C_ON, C_ON_BG = '2F7D4F', 'E7F3EC'
C_GRAY, C_GRAY_BG = '6B6B6B', 'F2F2F2'
C_RED = '9B2222'
C_YEL = 'FFD000'


# ---------- 글자 서식 ----------
# 한 줄은 조각(run) 목록. 조각 = (글자, 서식문자열). 서식: i 기울임, b 굵게, s 아래첨자, m 수식글꼴
def T(*parts):
    out = []
    for p in parts:
        out.append((p, '') if isinstance(p, str) else p)
    return out


def I(t):
    return (t, 'i')


def SUB(t):
    return (t, 's')


def ISUB(t):
    return (t, 'is')


def M(t):
    return (t, 'm')


def fill_text(tf, lines, size, color='000000', align=PP_ALIGN.CENTER, bold=False, italic=False,
              anchor=MSO_ANCHOR.MIDDLE, spacing=1.0):
    tf.clear()
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for li, line in enumerate(lines):
        para = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = spacing
        for text, fmt in line:
            r = para.add_run()
            r.text = text
            f = r.font
            f.size = Pt(size)
            f.name = MATHFONT if 'm' in fmt else FONT
            f.bold = bold or ('b' in fmt)
            f.italic = italic or ('i' in fmt)
            f.color.rgb = RGBColor.from_string(color)
            if 's' in fmt:
                r.font._rPr.set('baseline', '-25000')


# ---------- 도형 ----------
class Canvas:
    """matplotlib 와 같은 좌표계 (인치, y 는 아래에서 위로) 를 pptx 좌표로 바꿔 그린다."""

    def __init__(self, slide, H):
        self.s = slide
        self.H = H

    def Y(self, y):  # 위쪽 기준 좌표
        return self.H - y

    def rbox(self, x, y0, w, y1, ec, fc='FFFFFF', lw=0.7, radius=0.04, dash=False, name=None):
        h = y1 - y0
        sh = self.s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(self.Y(y1)),
                                     Inches(w), Inches(h))
        sh.adjustments[0] = min(0.5, radius / min(w, h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fc)
        sh.line.color.rgb = RGBColor.from_string(ec)
        sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = 4
        sh.shadow.inherit = False
        if name:
            sh.name = name
        return sh

    def box(self, x, y0, w, y1, lines, ec, fc='FFFFFF', size=7.0, color='000000', italic=False,
            name=None, lw=0.7):
        sh = self.rbox(x, y0, w, y1, ec, fc, lw=lw, name=name)
        fill_text(sh.text_frame, lines, size, color=color, italic=italic, spacing=1.05)
        return sh

    def text(self, x, y, lines, size=6.5, color='000000', ha='center', va='center', bold=False,
             italic=False, w=None, h=None, name=None):
        nlines = len(lines)
        h = h or (size / 72.0) * 1.25 * nlines + 0.02
        w = w or 2.5
        if ha == 'center':
            left = x - w / 2
        elif ha == 'left':
            left = x
        else:
            left = x - w
        if va == 'center':
            top = self.Y(y) - h / 2
        elif va == 'bottom':
            top = self.Y(y) - h
        else:
            top = self.Y(y)
        tb = self.s.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
        align = {'center': PP_ALIGN.CENTER, 'left': PP_ALIGN.LEFT, 'right': PP_ALIGN.RIGHT}[ha]
        fill_text(tb.text_frame, lines, size, color=color, align=align, bold=bold, italic=italic,
                  spacing=1.05)
        if name:
            tb.name = name
        return tb

    def _arrowhead(self, shape, color):
        ln = shape.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn('a:tailEnd'))
        tail.set('type', 'triangle')
        tail.set('w', 'sm')
        tail.set('len', 'sm')

    def arrow(self, p0, p1, color='000000', lw=0.9, head=True):
        c = self.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p0[0]), Inches(self.Y(p0[1])),
                                        Inches(p1[0]), Inches(self.Y(p1[1])))
        c.line.color.rgb = RGBColor.from_string(color)
        c.line.width = Pt(lw)
        if head:
            self._arrowhead(c, color)
        return c

    def poly(self, pts, color='000000', lw=0.9, head=True):
        """꺾은선 화살표 (마지막 점에 화살촉)."""
        e = [(Inches(x), Inches(self.Y(y))) for x, y in pts]
        fb = self.s.shapes.build_freeform(e[0][0], e[0][1], scale=1.0)
        fb.add_line_segments(e[1:], close=False)
        sh = fb.convert_to_shape()
        sh.fill.background()
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(lw)
        sh.shadow.inherit = False
        if head:
            self._arrowhead(sh, color)
        return sh

    def ellipse(self, cx, cy, a, b, color, alpha):
        sh = self.s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - a / 2), Inches(self.Y(cy) - b / 2),
                                     Inches(a), Inches(b))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(color)
        clr = sh.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
        al = etree.SubElement(clr, qn('a:alpha'))
        al.set('val', str(int(alpha * 100000)))
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(0.6)
        sh.shadow.inherit = False
        return sh


def new_pres(W, H):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 빈 레이아웃
    return prs, slide


# ======================= Fig. 1 =======================
def make_fig1():
    W, H = 7.16, 2.85
    prs, slide = new_pres(W, H)
    c = Canvas(slide, H)
    YB, YT = 0.42, 2.52

    def panel(x, w, title, ec, fc):
        c.rbox(x, YB, w, YT, ec, fc, lw=0.9, radius=0.06)
        c.text(x + w / 2, YT - 0.12, [T(title)], size=7.6, color=ec, bold=True, w=w)

    P1x, P1w = 0.05, 1.62
    P2x, P2w = 1.93, 2.20
    P3x, P3w = 4.39, 1.17
    P4x, P4w = 5.82, 1.29

    # (1) 선박
    panel(P1x, P1w, 'Vessel with gear shift dwell', C_PLANT, C_PLANT_BG)
    bx, bw = P1x + 0.09, P1w - 0.18
    c.box(bx, 1.86, bw, 2.26, [T('Gear shift ', I('g'), ' ∈ {R, N, F}'),
                               T('held ', I('D'), ISUB('g'), ' = 6 steps (3 s)')], C_PLANT)
    c.box(bx, 1.36, bw, 1.76, [T('Sterndrive steering ', I('δ')), T('rate limited, no dwell')], C_PLANT)
    c.box(bx, 0.86, bw, 1.26, [T('Disturbance ', I('d'), ' ∈ ', M('𝒲')), T('wind and current')], C_PLANT)
    c.text(P1x + P1w / 2, 0.64, [T('Locked in N, thrust and'), T('steering force both vanish')],
           size=6.8, color=C_RED, italic=True, w=P1w)

    # (2) Stage 1
    panel(P2x, P2w, 'Stage 1 offline, dwell indexed RIS', C_OFF, C_OFF_BG)
    bx, bw = P2x + 0.09, P2w - 0.18
    c.box(bx, 1.90, bw, 2.26, [T('Augmented state'),
                               T(I('ξ'), ' = (', I('x̃'), ', ', I('ỹ'), ', ', I('ψ̃'), ', ', I('u'), ', ',
                                 I('v'), ', ', I('r'), ', ', I('δ'), ', ', I('g'), ', ', I('τ'), ISUB('g'), ')')],
          C_OFF)
    half = (bw - 0.08) / 2
    c.box(bx, 1.42, half, 1.80, [T('Protagonist ', I('π'), ISUB('h')), T('max over ', I('C'), '(', I('ξ'), ')')], C_OFF)
    c.box(bx + half + 0.08, 1.42, half, 1.80, [T('Adversary ', I('μ')), T('min over ', M('𝒲'))], C_OFF)
    c.box(bx, 0.94, bw, 1.32, [T('Twin safety critics ', I('Q'), ISUB('h,1'), ', ', I('Q'), ISUB('h,2')),
                               T('one step discounted backup, masked')], C_OFF)
    cx, cy = bx + 0.40, 0.64
    for a, b, al in [(0.74, 0.34, 0.16), (0.50, 0.23, 0.30), (0.26, 0.12, 0.55)]:
        c.ellipse(cx, cy, a, b, C_OFF, al)
    for dx, lab, col in [(0.31, '0', C_OFF), (0.19, '3', C_OFF), (0.0, '5', 'FFFFFF')]:
        c.text(cx + dx, cy, [T(lab)], size=6.2, color=col, w=0.12)
    c.text(cx, cy + 0.235, [T(I('τ'), ISUB('g'), ' = 0, 3, 5')], size=6.3, color=C_OFF, w=0.8)
    c.text(bx + bw - 0.02, 0.67, [T('Safe set family'), T(M('𝒮'), '(', I('g'), ', ', I('τ'), ISUB('g'), '), smaller'),
                                  T('deeper in a lock')], size=6.8, ha='right', w=1.1)

    # (3) Stage 2
    panel(P3x, P3w, 'Stage 2 offline', C_OFF, C_OFF_BG)
    bx, bw = P3x + 0.08, P3w - 0.16
    c.box(bx, 1.70, bw, 2.24, [T('Task policy ', I('π'), ISUB('t')), T('soft actor critic')], C_OFF)
    c.box(bx, 1.06, bw, 1.58, [T('Lagrangian on'),
                               T(I('q'), ' ≥ ', I('τ'), ', ', I('λ'), ' ≥ ', I('λ'), SUB('min'))], C_OFF)
    c.text(P3x + P3w / 2, 0.74, [T('safety critic ', I('q')), T('frozen from'), T('Stage 1')],
           size=6.6, color=C_GRAY, w=P3w)

    # (4) 온보드
    panel(P4x, P4w, 'Onboard, every 0.5 s', C_ON, C_ON_BG)
    bx, bw = P4x + 0.08, P4w - 0.16
    c.box(bx, 1.88, bw, 2.26, [T('EKF state ', I('ξ'), ISUB('k')), T('at 10 Hz')], C_ON)
    c.box(bx, 1.16, bw, 1.76, [T('Shield'),
                               T('apply ', I('a'), ISUB('t'), ' if ', I('q'), '(', I('ξ'), ', ', I('a'), ISUB('t'), ') ≥ ', I('τ')),
                               T('else apply ', I('π'), ISUB('h'), '(', I('ξ'), ')')], C_ON)
    c.box(bx, 0.52, bw, 1.04, [T('PLC and'), T('servo cylinders')], C_ON)
    c.arrow((P4x + P4w / 2, 1.88), (P4x + P4w / 2, 1.76), color=C_ON)
    c.arrow((P4x + P4w / 2, 1.16), (P4x + P4w / 2, 1.04), color=C_ON)

    # 패널 사이 화살표
    ym = 1.56
    c.arrow((P1x + P1w, ym), (P2x, ym))
    c.text((P1x + P1w + P2x) / 2, ym + 0.13, [T('model')], size=6.5, w=0.4)
    c.arrow((P2x + P2w, ym), (P3x, ym))
    c.text((P2x + P2w + P3x) / 2, ym + 0.13, [T(I('q'), ', ', I('π'), ISUB('h'))], size=6.5, w=0.4)
    c.arrow((P3x + P3w, ym), (P4x, ym))
    c.text((P3x + P3w + P4x) / 2, ym + 0.13, [T(I('π'), ISUB('t'))], size=6.5, w=0.4)

    # 위쪽 경로
    yt = 2.68
    xs, xe = P2x + P2w * 0.75, P4x + P4w * 0.30
    c.poly([(xs, YT), (xs, yt), (xe, yt), (xe, YT)])
    c.text((xs + xe) / 2, yt + 0.09, [T('certified ', I('q'), ', ', I('π'), ISUB('h'), ', and margin ', I('τ'))],
           size=6.5, w=2.4)

    # 아래쪽 폐루프
    y1, y2 = 0.27, 0.11
    x_on, x_pl = P4x + P4w * 0.5, P1x + P1w * 0.5
    c.poly([(x_on, YB), (x_on, y1), (x_pl + 0.15, y1), (x_pl + 0.15, YB)], color=C_ON)
    c.text((x_on + x_pl) / 2, y1 + 0.08, [T('gear and steering commands to the vessel')], size=6.5,
           color=C_ON, w=3.0)
    c.poly([(x_pl - 0.15, YB), (x_pl - 0.15, y2), (x_on + 0.25, y2), (x_on + 0.25, YB)], color=C_PLANT)
    c.text((x_on + x_pl) / 2, y2 - 0.075, [T('GNSS and IMU measurements')], size=6.5, color=C_PLANT, w=3.0)

    out = os.path.join(OUTDIR, 'fig1_framework.pptx')
    prs.save(out)
    return out


# ======================= Fig. 2 =======================
def make_fig2():
    W, H = 7.16, 2.45
    prs, slide = new_pres(W, H)
    c = Canvas(slide, H)
    tmp = os.path.join(HERE, '_crop')
    os.makedirs(tmp, exist_ok=True)

    # (a) 선박 사진
    im = Image.open(os.path.join(PH, 'vessel_side.jpg')).crop((140, 300, 1340, 1040))
    pa = os.path.join(tmp, 'a.jpg')
    im.save(pa, quality=92)
    wa = 2.62
    ha = wa * im.height / im.width
    ya = 0.26 + (2.05 - ha) / 2
    pic = slide.shapes.add_picture(pa, Inches(0.02), Inches(c.Y(ya + ha)), Inches(wa), Inches(ha))
    pic.name = 'photo_a_vessel'

    # (b) 선미 기관실 사진
    x0, y0, x1, y1 = 300, 360, 760, 960
    im = Image.open(os.path.join(PH, 'stern_bay.jpg')).crop((x0, y0, x1, y1))
    pb = os.path.join(tmp, 'b.jpg')
    im.save(pb, quality=92)
    hb = 2.05
    wb = hb * im.width / im.height
    xb = 0.02 + wa + 0.10
    pic = slide.shapes.add_picture(pb, Inches(xb), Inches(c.Y(0.26 + hb)), Inches(wb), Inches(hb))
    pic.name = 'photo_b_stern_bay'

    def px(ix, iy):  # 잘린 사진 픽셀 → 그림 좌표 (y 아래에서 위로)
        return xb + ix / im.width * wb, 0.26 + hb - iy / im.height * hb

    targets = {'shift': (638, 691), 'steer': (413, 539), 'pivot': (485, 858)}
    labels = [('steer', ['Steering servo', 'cylinder'], (120, 40)),
              ('shift', ['Shift servo', 'cylinder (R/N/F)'], (345, 150)),
              ('pivot', ['Drive steering', 'pivot'], (150, 560))]
    for key, lines, (lx, ly) in labels:
        tx, ty = targets[key]
        p_t = px(tx - x0, ty - y0)
        p_l = px(lx, ly)
        bw_, bh_ = 0.78, 0.24
        # 화살표를 먼저 그려 라벨 상자 아래에 깔리게 함
        c.arrow((p_l[0], p_l[1] - bh_ / 2 if p_t[1] < p_l[1] else p_l[1] + bh_ / 2), p_t, color=C_YEL, lw=1.0)
        c.box(p_l[0] - bw_ / 2, p_l[1] - bh_ / 2, bw_, p_l[1] + bh_ / 2, [T(t) for t in lines], C_PLANT,
              size=6.4, lw=0.6)

    # (c) 하드웨어 연결도
    xc = xb + wb + 0.12
    wc = W - xc - 0.02
    oy = 0.26  # 패널 아래쪽 기준

    def B(x, y, w, h, lines, ec, fc, size=6.6):
        c.box(xc + x, oy + y, w, oy + y + h, lines, ec, fc, size=size)

    def L(p0, p1):
        c.arrow((xc + p0[0], oy + p0[1]), (xc + p1[0], oy + p1[1]), lw=0.8)

    def LT(x, y, s, ha='center'):
        c.text(xc + x, oy + y, [T(s)], size=5.9, color=C_GRAY, ha=ha, w=0.6)

    bw2 = (wc - 0.12) / 2
    B(0.0, 1.62, bw2, 0.38, [T('Here3 GNSS'), T('receiver')], C_PLANT, C_PLANT_BG)
    B(bw2 + 0.12, 1.62, bw2, 0.38, [T('Pixhawk Cube'), T('with IMU')], C_PLANT, C_PLANT_BG)
    L((bw2, 1.81), (bw2 + 0.12, 1.81))
    B(0.0, 1.00, wc, 0.42, [T('Onboard computer, ROS 2'), T('EKF at 10 Hz, shielded controller every 0.5 s')],
      C_ON, C_ON_BG)
    xm = bw2 + 0.12 + bw2 / 2
    L((xm, 1.62), (xm, 1.42))
    LT(xm + 0.05, 1.52, 'serial', ha='left')
    B(0.0, 0.58, wc, 0.26, [T('LS XGB PLC')], C_GRAY, C_GRAY_BG)
    L((wc / 2, 1.00), (wc / 2, 0.84))
    LT(wc / 2 + 0.05, 0.92, 'Ethernet', ha='left')
    aw = (wc - 0.16) / 3
    acts = [[T('Shift servo'), T('cylinder'), T('R/N/F')],
            [T('Steering servo'), T('cylinder'), T(I('δ'))],
            [T('Engine'), T('speed')]]
    for i, lines in enumerate(acts):
        B(i * (aw + 0.08), 0.0, aw, 0.40, lines, C_PLANT, C_PLANT_BG, size=6.2)
    for i, t in enumerate(['EtherCAT', 'EtherCAT', 'CAN']):
        cxx = i * (aw + 0.08) + aw / 2
        L((cxx, 0.58), (cxx, 0.40))
        LT(cxx + 0.04, 0.49, t, ha='left')

    # 패널 이름
    for xx, s in [(0.02 + wa / 2, '(a)'), (xb + wb / 2, '(b)'), (xc + wc / 2, '(c)')]:
        c.text(xx, 0.10, [T(s)], size=7.5, w=0.4)

    out = os.path.join(OUTDIR, 'fig2_vessel.pptx')
    prs.save(out)
    return out


if __name__ == '__main__':
    print(make_fig1())
    print(make_fig2())
