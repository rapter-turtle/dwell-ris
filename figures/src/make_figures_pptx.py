"""Fig. 1, 2, 3 을 손으로 고칠 수 있는 pptx 로 만든다 (모든 상자, 글자, 화살표가 개별 도형).

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
        no_effect(c)
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


class TC:
    """위쪽 왼쪽이 원점인 좌표계 (인치) 로 그리는 도구. Fig. 1 용."""

    def __init__(self, slide):
        self.s = slide

    def rect(self, x, y, w, h, ec, fc='FFFFFF', lw=0.7, radius=0.04, dash=False, name=None):
        sh = self.s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.adjustments[0] = min(0.5, radius / min(w, h))
        if fc is None:
            sh.fill.background()
        else:
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

    def node(self, x, y, w, h, lines, ec, fc='FFFFFF', size=7.0, color='000000', lw=0.7, bold=False, name=None):
        sh = self.rect(x, y, w, h, ec, fc, lw=lw, name=name)
        fill_text(sh.text_frame, lines, size, color=color, bold=bold, spacing=1.05)
        return sh

    def label(self, x, y, w, h, lines, size=6.5, color='000000', align='center', bold=False, italic=False,
              anchor=MSO_ANCHOR.MIDDLE, name=None):
        tb = self.s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        al = {'center': PP_ALIGN.CENTER, 'left': PP_ALIGN.LEFT, 'right': PP_ALIGN.RIGHT}[align]
        fill_text(tb.text_frame, lines, size, color=color, align=al, bold=bold, italic=italic,
                  anchor=anchor, spacing=1.05)
        if name:
            tb.name = name
        return tb

    def _end(self, shape, tag):
        ln = shape.line._get_or_add_ln()
        e = etree.SubElement(ln, qn(tag))
        e.set('type', 'triangle')
        e.set('w', 'sm')
        e.set('len', 'sm')

    def arrow(self, p0, p1, color='000000', lw=0.9, both=False, head=True, dash=False):
        c = self.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p0[0]), Inches(p0[1]),
                                        Inches(p1[0]), Inches(p1[1]))
        c.line.color.rgb = RGBColor.from_string(color)
        c.line.width = Pt(lw)
        if dash:
            c.line.dash_style = 4
        if both:
            self._end(c, 'a:headEnd')
        if head:
            self._end(c, 'a:tailEnd')
        no_effect(c)
        return c

    def poly(self, pts, color='000000', lw=0.9, head=True):
        e = [(Inches(x), Inches(y)) for x, y in pts]
        fb = self.s.shapes.build_freeform(e[0][0], e[0][1], scale=1.0)
        fb.add_line_segments(e[1:], close=False)
        sh = fb.convert_to_shape()
        sh.fill.background()
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(lw)
        sh.shadow.inherit = False
        if head:
            self._end(sh, 'a:tailEnd')
        return sh

    def polygon(self, pts, fc, ec=None, lw=0.6, name=None):
        e = [(Inches(x), Inches(y)) for x, y in pts]
        fb = self.s.shapes.build_freeform(e[0][0], e[0][1], scale=1.0)
        fb.add_line_segments(e[1:], close=True)
        sh = fb.convert_to_shape()
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fc)
        if ec:
            sh.line.color.rgb = RGBColor.from_string(ec)
            sh.line.width = Pt(lw)
        else:
            sh.line.fill.background()
        sh.shadow.inherit = False
        if name:
            sh.name = name
        return sh

    def ellipse(self, cx, cy, a, b, color, alpha):
        sh = self.s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - a / 2), Inches(cy - b / 2), Inches(a), Inches(b))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(color)
        clr = sh.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
        etree.SubElement(clr, qn('a:alpha')).set('val', str(int(alpha * 100000)))
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(0.6)
        sh.shadow.inherit = False
        return sh

    def boat(self, x, y, L, B, color=C_PLANT):
        """위에서 본 선체 (선미 왼쪽, 선수 오른쪽) 와 선미 드라이브."""
        hull = [(x, y + 0.12 * B), (x + 0.62 * L, y), (x + 0.86 * L, y + 0.12 * B), (x + L, y + 0.5 * B),
                (x + 0.86 * L, y + 0.88 * B), (x + 0.62 * L, y + B), (x, y + 0.88 * B)]
        self.polygon(hull, color, name='boat_hull')
        self.polygon([(x - 0.07 * L, y + 0.38 * B), (x, y + 0.38 * B), (x, y + 0.62 * B),
                      (x - 0.07 * L, y + 0.62 * B)], '1B3A5C', name='boat_sterndrive')

    def gusts(self, x, y, n=3, dy=0.07, L=0.16, color='8A8A8A'):
        for i in range(n):
            self.arrow((x, y + i * dy), (x + L, y + i * dy), color=color, lw=0.8)

    def nn(self, x, y, w, h, color, layers=(3, 4, 3), frozen=False, name=None):
        """작은 신경망 그림 (테두리, 노드, 연결선) 을 한 그룹으로 만든다."""
        grp = self.s.shapes.add_group_shape()
        shp = grp.shapes
        node_c = C_GRAY if frozen else color
        edge_c = 'C8C8C8' if frozen else 'E3B48A'
        border = shp.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        border.adjustments[0] = 0.15
        border.fill.solid()
        border.fill.fore_color.rgb = RGBColor.from_string(C_GRAY_BG if frozen else 'FFFFFF')
        border.line.color.rgb = RGBColor.from_string(node_c)
        border.line.width = Pt(0.7)
        if frozen:
            border.line.dash_style = 4
        border.shadow.inherit = False
        nmax = max(layers)
        dd = min(0.055, (h - 0.06) / (nmax * 1.7))
        xs = [x + 0.07 + i * (w - 0.14) / (len(layers) - 1) for i in range(len(layers))]
        # 층마다 노드를 같은 간격으로 가운데 정렬
        pts = []
        for xi, n in zip(xs, layers):
            span = (h - 0.10) * (n - 1) / max(nmax - 1, 1)
            pts.append([(xi, y + h / 2 - span / 2 + (span / (n - 1) * j if n > 1 else 0)) for j in range(n)])
        for a, b in zip(pts[:-1], pts[1:]):
            for p in a:
                for q in b:
                    c = shp.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(p[0]), Inches(p[1]), Inches(q[0]), Inches(q[1]))
                    c.line.color.rgb = RGBColor.from_string(edge_c)
                    c.line.width = Pt(0.35)
                    no_effect(c)
        for col in pts:
            for (px_, py_) in col:
                o = shp.add_shape(MSO_SHAPE.OVAL, Inches(px_ - dd / 2), Inches(py_ - dd / 2), Inches(dd), Inches(dd))
                o.fill.solid()
                o.fill.fore_color.rgb = RGBColor.from_string(node_c)
                o.line.fill.background()
                o.shadow.inherit = False
        if name:
            grp.name = name
        return grp

    def cylinder(self, x, y, w, h, lines, ec, fc, size=6.2, name=None):
        sh = self.s.shapes.add_shape(MSO_SHAPE.CAN, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fc)
        sh.line.color.rgb = RGBColor.from_string(ec)
        sh.line.width = Pt(0.7)
        sh.shadow.inherit = False
        fill_text(sh.text_frame, lines, size, spacing=1.0)
        if name:
            sh.name = name
        return sh

    def picture(self, path, x, y, w=None, h=None, name=None):
        im = Image.open(path)
        if w is not None and h is None:
            h = w * im.height / im.width
        elif h is not None and w is None:
            w = h * im.width / im.height
        p = self.s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            p.name = name
        return x, y, w, h


def no_effect(shape):
    """선 도형에 '효과 없음' 을 직접 지정해 테마 그림자를 끈다."""
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None:
        etree.SubElement(spPr, qn('a:effectLst'))


def _save(prs, out):
    """저장한 뒤 테마의 효과 스타일에서 그림자를 지운다 (복사해 간 도형에도 그림자가 생기지 않게)."""
    import re
    import shutil
    import zipfile
    prs.save(out)
    tmp = out + '.tmp'
    with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/theme/') and item.filename.endswith('.xml'):
                x = data.decode('utf-8')
                x = re.sub(r'<a:effectStyleLst>.*?</a:effectStyleLst>',
                           '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 +
                           '</a:effectStyleLst>', x, flags=re.S)
                data = x.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, out)


def new_pres(W, H):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 빈 레이아웃
    return prs, slide


# ======================= Fig. 1 =======================
def make_fig1():
    """왼쪽 Offline (병렬 시뮬레이터, replay buffer, 신경망으로 표현한 적대적 RL), 오른쪽 Online (실제 선박 사진)."""
    W, H = 7.16, 2.45
    prs, slide = new_pres(W, H)
    t = TC(slide)

    OX, OY, OW, OH = 0.04, 0.04, 4.50, 2.37
    NX, NY, NW, NH = 4.94, 0.04, 2.18, 2.37
    t.rect(OX, OY, OW, OH, C_OFF, C_OFF_BG, lw=1.0, radius=0.08, name='region_offline')
    t.rect(NX, NY, NW, NH, C_ON, C_ON_BG, lw=1.0, radius=0.08, name='region_online')
    t.label(OX + 0.10, OY + 0.05, 3.0, 0.22, [T(('Offline', 'b'), ' RL training in simulation')],
            size=8.0, color=C_OFF, align='left')
    t.label(NX + 0.10, NY + 0.05, 2.0, 0.22, [T(('Online', 'b'), ' deployment on the vessel')],
            size=8.0, color=C_ON, align='left')

    # ----- 병렬 시뮬레이터 -----
    t.label(0.12, 0.36, 1.06, 0.16, [T(('Simulators ', 'b'), '×', I('N'))], size=7.0, color=C_PLANT)
    for k, (dx, fc) in enumerate([(0.12, 'D5E2EF'), (0.06, 'E8EFF6'), (0.0, 'FFFFFF')]):
        t.rect(0.14 + dx, 0.58 + dx, 0.86, 0.78, C_PLANT, fc, lw=0.7, radius=0.05, name=f'sim_card_{k}')
    t.gusts(0.22, 0.70, n=3, dy=0.055, L=0.14)
    t.label(0.40, 0.69, 0.5, 0.14, [T(I('d'), ' ∈ ', M('𝒲'))], size=6.4, color=C_GRAY, align='left')
    t.boat(0.30, 0.90, 0.58, 0.20)
    t.label(0.14, 1.17, 0.86, 0.14, [T('dwell ', I('D'), ISUB('g'), ' masked')], size=6.2, color=C_PLANT)
    t.label(0.10, 1.56, 1.10, 0.52, [T('gear ', I('g'), ' ∈ {R, N, F}'), T('steering ', I('δ')),
                                     T('wind and current ', I('d'))], size=6.4)

    # ----- Stage 1 -----
    s1x, s1y, s1w, s1h = 1.36, 0.34, 3.08, 1.10
    t.rect(s1x, s1y, s1w, s1h, C_OFF, 'FFFFFF', name='stage1')
    t.label(s1x + 0.08, s1y + 0.03, 2.9, 0.18, [T(('Stage 1', 'b'), '  dwell constrained RIS by adversarial RL')],
            size=7.0, color=C_OFF, align='left')
    t.cylinder(s1x + 0.10, s1y + 0.32, 0.44, 0.52, [T('replay'), T('buffer')], C_OFF, C_OFF_BG, size=6.0,
               name='replay_buffer')
    # twin critics
    t.nn(s1x + 0.78, s1y + 0.30, 0.56, 0.42, C_OFF, name='critic_2')
    t.nn(s1x + 0.72, s1y + 0.36, 0.56, 0.42, C_OFF, name='critic_1')
    t.label(s1x + 0.62, s1y + 0.79, 0.80, 0.14, [T(I('Q'), ISUB('h,1'), ', ', I('Q'), ISUB('h,2'))], size=6.5)
    t.arrow((s1x + 0.54, s1y + 0.58), (s1x + 0.72, s1y + 0.58), color=C_OFF)
    # protagonist, adversary
    t.nn(s1x + 1.62, s1y + 0.26, 0.44, 0.30, C_OFF, layers=(3, 3, 2), name='protagonist')
    t.nn(s1x + 1.62, s1y + 0.64, 0.44, 0.30, C_OFF, layers=(3, 3, 2), name='adversary')
    t.label(s1x + 2.08, s1y + 0.25, 0.42, 0.32, [T(I('π'), ISUB('h')), T(('max', ''))], size=6.5, align='left')
    t.label(s1x + 2.08, s1y + 0.63, 0.42, 0.32, [T(I('μ')), T('min')], size=6.5, align='left')
    t.arrow((s1x + 1.34, s1y + 0.52), (s1x + 1.62, s1y + 0.41), color=C_OFF, dash=True)
    t.arrow((s1x + 1.34, s1y + 0.64), (s1x + 1.62, s1y + 0.79), color=C_OFF, dash=True)
    # 안전집합 족
    ex, ey = s1x + 2.70, s1y + 0.60
    for a, b, al in [(0.62, 0.40, 0.16), (0.42, 0.27, 0.30), (0.22, 0.14, 0.55)]:
        t.ellipse(ex, ey, a, b, C_OFF, al)
    for dx, lab, col in [(0.26, '0', C_OFF), (0.16, '3', C_OFF), (0.0, '5', 'FFFFFF')]:
        t.label(ex + dx - 0.05, ey - 0.07, 0.10, 0.14, [T(lab)], size=5.8, color=col)
    t.label(ex - 0.40, s1y + 0.24, 0.80, 0.14, [T(M('𝒮'), '(', I('g'), ', ', I('τ'), ISUB('g'), ')')],
            size=6.5, color=C_OFF)
    # 1-step 목표
    t.label(s1x + 0.10, s1y + 0.93, 2.9, 0.15,
            [T('one step target  ', I('y'), ' = (1 − ', I('γ'), ISUB('h'), ') ', I('h'), ' + ', I('γ'), ISUB('h'),
               ' min{', I('h'), ', ', I('Q̄'), ISUB('h'), '(', I('ξ'), '′)}')], size=6.4, color=C_OFF)

    # ----- Stage 2 -----
    s2x, s2y, s2w, s2h = 1.36, 1.52, 3.08, 0.80
    t.rect(s2x, s2y, s2w, s2h, C_OFF, 'FFFFFF', name='stage2')
    t.label(s2x + 0.08, s2y + 0.03, 2.0, 0.18, [T(('Stage 2', 'b'), '  task policy by SAC')],
            size=7.0, color=C_OFF, align='left')
    t.label(s2x + 0.10, s2y + 0.30, 1.0, 0.36, [T('maximize reward'), T('s.t. ', I('q'), ' ≥ ', I('τ'))], size=6.5)
    t.nn(s2x + 1.32, s2y + 0.28, 0.50, 0.36, C_OFF, name='task_policy')
    t.label(s2x + 1.22, s2y + 0.65, 0.70, 0.14, [T(I('π'), ISUB('t'))], size=6.8)
    t.nn(s2x + 2.36, s2y + 0.28, 0.50, 0.36, C_OFF, frozen=True, name='critic_frozen')
    t.label(s2x + 2.21, s2y + 0.65, 0.80, 0.14, [T(I('Q'), ISUB('h'), ' frozen')], size=6.5, color=C_GRAY)
    t.arrow((s2x + 2.36, s2y + 0.46), (s2x + 1.82, s2y + 0.46), color=C_OFF)
    t.label(s2x + 1.82, s2y + 0.27, 0.54, 0.16, [T(I('λ'), ' ≥ ', I('λ'), SUB('min'))], size=6.2)
    # Stage 1 에서 고정된 critic 으로
    t.arrow((ex, s1y + s1h), (ex, s2y + 0.28), color=C_OFF)
    t.label(ex + 0.04, s2y + 0.04, 0.5, 0.14, [T('freeze')], size=6.0, color=C_GRAY, align='left')

    # 시뮬레이터 왕복
    for yy, top, bot in [(s1y + 0.58, 'c, d', 'ξ′, h'), (s2y + 0.40, 'a', 'ξ′, r')]:
        t.arrow((1.12, yy), (s1x, yy), color=C_PLANT, both=True)
        t.label(1.10, yy - 0.17, 0.28, 0.14, [T(I(top))], size=6.0, color=C_PLANT)
        t.label(1.10, yy + 0.02, 0.28, 0.14, [T(I(bot))], size=6.0, color=C_PLANT)

    # ----- Offline 에서 Online 으로 -----
    ym = 1.24
    t.arrow((OX + OW, ym), (NX, ym), color='000000', lw=1.4)
    t.label(OX + OW - 0.02, ym - 0.24, 0.44, 0.16, [T('deploy')], size=6.5)
    t.label(OX + OW - 0.02, ym + 0.06, 0.44, 0.16, [T('frozen')], size=6.5, color=C_GRAY)

    # ----- Online 안쪽 -----
    t.node(NX + 0.12, 0.42, 0.50, 0.30, [T('EKF')], C_ON, size=7.0)
    t.nn(NX + 0.80, 0.40, 0.46, 0.32, C_OFF, name='task_policy_onboard')
    t.label(NX + 0.80, 0.73, 0.46, 0.13, [T(I('π'), ISUB('t'))], size=6.6)
    t.node(NX + 1.42, 0.36, 0.66, 0.50, [T('Shield'), T(I('q'), ' ≥ ', I('τ')), T('else ', I('π'), ISUB('h'))],
           C_OFF, size=6.6)
    t.arrow((NX + 0.62, 0.57), (NX + 0.80, 0.57), color=C_ON)
    t.arrow((NX + 1.26, 0.57), (NX + 1.42, 0.57), color=C_ON)
    # 실제 선박 사진
    im = Image.open(os.path.join(PH, 'vessel_side.jpg')).crop((60, 300, 1260, 1040))
    pth = os.path.join(PH, 'crops', 'fig1_vessel.jpg')
    os.makedirs(os.path.dirname(pth), exist_ok=True)
    im.save(pth, quality=92)
    px, py, pw, ph = t.picture(pth, NX + 0.11, 1.12, w=1.96, name='photo_vessel')
    t.gusts(px + 0.08, py + 0.10, n=3, dy=0.06, L=0.18, color='FFFFFF')
    t.label(px + 0.30, py + 0.10, 0.3, 0.14, [T(I('d'))], size=6.8, color='FFFFFF', align='left')
    # 명령과 측정
    xs_ = NX + 1.75
    t.arrow((xs_, 0.86), (xs_, py), color=C_ON)
    t.label(xs_ + 0.04, 0.92, 0.36, 0.14, [T(I('g'), ', ', I('δ'))], size=6.6, align='left')
    xm_ = NX + 0.37
    t.arrow((xm_, py), (xm_, 0.72), color=C_ON)
    t.label(xm_ + 0.04, 0.92, 0.80, 0.14, [T('GPS, AHRS')], size=6.0, color=C_GRAY, align='left')

    out = os.path.join(OUTDIR, 'fig1_framework.pptx')
    _save(prs, out)
    return out


# ======================= Fig. 2 =======================
C_LINK = '00A0E0'   # 연결선 (원본 5번 슬라이드의 하늘색)
C_PROTO = '1F6FB2'  # 통신 방식 글자
C_BOUND = 'D62828'  # 선박 제어부 점선
DECK = os.path.join(HERE, '..', '..', '..', 'Ship configuration', 'Ship mechanical system.pptx')


def _media(name):
    """원본 'Ship mechanical system.pptx' 의 그림 파일을 읽는다."""
    import io
    import zipfile
    z = zipfile.ZipFile(DECK)
    return Image.open(io.BytesIO(z.read('ppt/media/' + name))).convert('RGB')


def _crop_frac(im, l, t, r, b):
    w, h = im.size
    return im.crop((int(l * w), int(t * h), int(w - r * w), int(h - b * h)))


def _green_bbox(im, pad=4):
    """CAN 컨버터 사진에서 초록 기판 영역만 남긴다 (둘레의 한글 주석 제거)."""
    px = im.load()
    w, h = im.size
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            if g > 90 and g > r + 25 and g > b + 25:
                xs.append(x)
                ys.append(y)
    return im.crop((max(0, min(xs) - pad), max(0, min(ys) - pad), min(w, max(xs) + pad), min(h, max(ys) + pad)))


def _remove_blue_label(im, color=(42, 42, 46)):
    """사진의 파란 로고 스티커 영역을 하우징과 같은 짙은 회색으로 덮는다."""
    import cv2
    import numpy as np
    a = np.asarray(im).astype(np.int16)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    mask = ((b > 110) & (b > r + 40) & (b > g + 15)).astype(np.uint8)
    n, lab, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    if n <= 1:
        return im
    k = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    x, y, w, h = stats[k, :4]
    x0, y0 = max(0, x - 2), max(0, y - 2)
    x1, y1 = min(a.shape[1], x + w + 2), min(a.shape[0], y + h + 2)
    out = np.asarray(im).copy()
    rng = np.random.default_rng(0)
    patch = np.empty((y1 - y0, x1 - x0, 3), dtype=np.float32)
    grad = np.linspace(-6, 6, y1 - y0)[:, None]
    for c in range(3):
        patch[..., c] = color[c] + grad + rng.normal(0, 3, (y1 - y0, x1 - x0))
    out[y0:y1, x0:x1] = np.clip(patch, 0, 255).astype(np.uint8)
    return Image.fromarray(out)


def make_fig2():
    """(a) 실험선과 센서 (사용자 합성 사진), (b) 제어 하드웨어 구성 (원본 5번 슬라이드 형식)."""
    W, H = 7.16, 1.96
    prs, slide = new_pres(W, H)
    t = TC(slide)
    tmp = os.path.join(PH, 'crops')
    os.makedirs(tmp, exist_ok=True)

    def pic(im, fname, x, y, w=None, h=None, name=None):
        path = os.path.join(tmp, fname)
        im.save(path, quality=92)
        return t.picture(path, x, y, w=w, h=h, name=name)

    # ---------- (a) 실험선과 센서 ----------
    wa = 3.62
    _, _, _, ha = t.picture(os.path.join(PH, 'vessel_sensors.jpg'), 0.04, 0.04, w=wa, name='photo_a_vessel_sensors')
    ya_bottom = 0.04 + ha

    # ---------- (b) 제어 하드웨어 ----------
    b0 = 3.76
    bw = W - b0 - 0.04
    oy = -0.42  # Control PC 를 뺀 만큼 위로 올림
    t.rect(b0 + 0.02, 0.62 + oy, bw - 0.02, 1.46, C_BOUND, None, lw=1.2, radius=0.18, dash=True,
           name='boat_control_boundary')
    t.label(b0 + bw - 1.62, 0.64 + oy, 1.56, 0.13, [T('Boat control hardware')], size=6.4, color=C_BOUND,
            bold=True, align='right')
    # PLC
    lx_, ly_, lw2, lh2 = pic(_media('image15.jpeg'), 'plc.jpg', b0 + 0.14, 0.98 + oy, w=0.52, name='plc')
    t.rect(lx_, ly_, lw2, lh2, C_BOUND, None, lw=1.0, radius=0.01)
    t.label(lx_ - 0.02, ly_ + lh2 + 0.02, lw2 + 0.12, 0.13, [T(('LS XGB PLC', 'b'))], size=6.4)
    # (a) 사진 속 Control PC (노트북 파란 상자의 오른쪽 변) 에서 PLC 로 Ethernet
    sa = wa / 1400.0  # 축소본 1400 폭 기준 좌표를 그림 좌표로
    lap_x = 0.04 + 1258 * sa          # 노트북 파란 상자의 오른쪽 변
    y_lap = 0.04 + 470 * sa           # 그 변의 위쪽 부분
    y_eth = ly_ + lh2 / 2
    x_elb = b0 + 0.08                 # 점선 영역 안쪽에서 꺾음
    t.poly([(lap_x, y_lap), (x_elb, y_lap), (x_elb, y_eth), (lx_, y_eth)], color=C_LINK, lw=1.3)
    t.node(lap_x + 0.03, y_lap + 0.03, 0.42, 0.13, [T('Ethernet')], C_LINK, size=6.0, color=C_PROTO, lw=0.5)
    # 버스
    xbus = b0 + 1.02
    t.arrow((lx_ + lw2, ly_ + lh2 / 2), (xbus, ly_ + lh2 / 2), color=C_LINK, lw=1.3, head=False)
    # CAN 컨버터
    cx_, cy_, cw_, ch_ = pic(_green_bbox(_media('image16.png')), 'can.jpg', b0 + 1.20, 0.72 + oy, w=0.46,
                             name='can_converter')
    t.label(cx_ - 0.15, cy_ + ch_ + 0.02, cw_ + 0.30, 0.13, [T(('CAN converter', 'b'))], size=6.4)
    yc = cy_ + ch_ / 2
    # 서보 실린더
    srv = _remove_blue_label(_crop_frac(_media('image17.jpeg'), 0.0645, 0.1363, 0.1306, 0.1351))
    sx_, sy_, sw_, sh_ = pic(srv, 'servo.jpg', b0 + 1.20, 1.42 + oy, w=0.80, name='servo_cylinder')
    t.label(sx_ - 0.1, sy_ + sh_ + 0.02, sw_ + 0.2, 0.13, [T(('Servo cylinders', 'b'))], size=6.4)
    ys = sy_ + sh_ / 2
    t.arrow((xbus, yc), (xbus, ys), color=C_LINK, lw=1.3, head=False)
    t.arrow((xbus, yc), (cx_, yc), color=C_LINK, lw=1.3)
    t.arrow((xbus, ys), (sx_, ys), color=C_LINK, lw=1.3)
    t.label(xbus - 0.33, yc - 0.15, 0.32, 0.13, [T('RS232C')], size=6.0, color=C_PROTO, align='right')
    t.label(xbus - 0.55, ys + 0.02, 0.52, 0.13, [T('EtherCAT')], size=6.0, color=C_PROTO, align='right')
    # 추진기 (실제 선미 드라이브 사진, 그늘 보정본)
    sdi = Image.open(os.path.join(PH, 'sterndrive_lifted.jpg'))
    sdi = sdi.crop((0, int(0.12 * sdi.height), int(0.95 * sdi.width), int(0.97 * sdi.height)))
    ex_, ey_, ew_, eh_ = pic(sdi, 'thruster.jpg', b0 + bw - 1.08, 0.96 + oy, w=0.98, name='sterndrive_thruster')
    xe1 = ex_ + ew_ * 0.62
    t.poly([(cx_ + cw_, yc), (xe1, yc), (xe1, ey_)], color=C_LINK, lw=1.3)
    t.label(cx_ + cw_ + 0.04, yc - 0.15, 0.4, 0.13, [T('CAN')], size=6.0, color=C_PROTO, align='left')
    t.label(xe1 - 0.80, yc - 0.17, 0.77, 0.13, [T(('Engine speed', 'b'))], size=6.4, align='right')
    xe2 = ex_ + ew_ * 0.40
    t.poly([(sx_ + sw_, ys + 0.10), (xe2, ys + 0.10), (xe2, ey_ + eh_)], color=C_LINK, lw=1.3)
    t.label(sx_ + sw_ + 0.06, ys + 0.13, 1.2, 0.13, [T(('Shift and steering', 'b'))], size=6.4, align='left')

    ylab = max(ya_bottom, 2.08 + oy) + 0.04
    for xx, s_ in [(0.04 + wa / 2, '(a)'), (b0 + bw / 2, '(b)')]:
        t.label(xx - 0.2, ylab, 0.4, 0.14, [T(s_)], size=7.5)

    out = os.path.join(OUTDIR, 'fig2_vessel.pptx')
    _save(prs, out)
    return out


# ======================= Fig. 3 =======================
def make_fig3():
    """좌표계와 선미 드라이브 추력 도식 (단일 그림, 본문에서 0.6 단 폭)."""
    import math
    W, H = 1.86, 1.82
    prs, slide = new_pres(W, H)
    t = TC(slide)
    C_SET = C_ON

    def P(o, f, s, F, S):
        return (o[0] + f * F[0] + s * S[0], o[1] + f * F[1] + s * S[1])

    def dirv(theta_deg):  # 북쪽에서 시계방향 각도 → 슬라이드 좌표 단위벡터
        th = math.radians(theta_deg)
        return (math.sin(th), -math.cos(th))

    def arc(c, rho, th0, th1, n=14, color='000000', head=False, lw=0.7):
        pts = [(c[0] + rho * dirv(th0 + (th1 - th0) * i / (n - 1))[0],
                c[1] + rho * dirv(th0 + (th1 - th0) * i / (n - 1))[1]) for i in range(n)]
        return t.poly(pts, color=color, lw=lw, head=head)

    def ring(cx, cy, r, color, dash=True, lw=0.9):
        sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - r), Inches(cy - r), Inches(2 * r), Inches(2 * r))
        sh.fill.background()
        sh.line.color.rgb = RGBColor.from_string(color)
        sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = 4
        sh.shadow.inherit = False
        return sh

    # ---------- (a) 도식 ----------
    # 관성 좌표계 (x 북쪽 위, y 동쪽 오른쪽)
    Oi = (0.14, 1.62)
    t.arrow(Oi, (Oi[0], 1.24), lw=0.8)
    t.arrow(Oi, (0.52, Oi[1]), lw=0.8)
    t.label(Oi[0] - 0.12, 1.14, 0.24, 0.12, [T(I('x'))], size=6.8)
    t.label(0.52, Oi[1] - 0.07, 0.14, 0.14, [T(I('y'))], size=6.8, align='left')
    t.label(0.02, Oi[1] + 0.01, 0.2, 0.13, [T(I('O'), ISUB('i'))], size=6.4, align='left')
    # 유지 영역과 설정점
    C = (0.78, 0.74)
    Rr = 0.64
    ring(C[0], C[1], Rr, C_SET)
    t.polygon([(C[0] - 0.035, C[1]), (C[0], C[1] - 0.035), (C[0] + 0.035, C[1]), (C[0], C[1] + 0.035)], C_SET)
    t.label(C[0] - 0.46, C[1] - 0.17, 0.44, 0.13, [T('(', I('x'), ISUB('d'), ', ', I('y'), ISUB('d'), ')')],
            size=6.0, color=C_SET, align='right')
    rv = (C[0] - 0.60 * Rr, C[1] + 0.80 * Rr)
    t.arrow(C, rv, color=C_SET, lw=0.7)
    t.label((C[0] + rv[0]) / 2 - 0.19, (C[1] + rv[1]) / 2 - 0.07, 0.14, 0.14, [T(I('R'))], size=6.6, color=C_SET)
    # 선체
    psi = 35.0
    F = dirv(psi)
    S = dirv(psi + 90)
    Ob = (1.22, 0.64)
    hull_b = [(-0.27, -0.09), (0.10, -0.09), (0.20, -0.06), (0.27, 0.0), (0.20, 0.06), (0.10, 0.09), (-0.27, 0.09)]
    t.polygon([P(Ob, f, s, F, S) for f, s in hull_b], C_PLANT, name='hull_top_view')
    # 위치 오차
    t.arrow(C, Ob, color=C_GRAY, lw=0.7, dash=True)
    # 북쪽 기준선과 헤딩 ψ
    t.arrow(Ob, (Ob[0], Ob[1] - 0.40), color=C_GRAY, lw=0.6, head=False, dash=True)
    arc(Ob, 0.30, 0, psi, color='000000')
    pl = (Ob[0] + 0.40 * dirv(psi / 2)[0], Ob[1] + 0.40 * dirv(psi / 2)[1])
    t.label(pl[0] - 0.06, pl[1] - 0.07, 0.12, 0.14, [T(I('ψ'))], size=6.8)
    # 선체 좌표 u, v, r
    ue = P(Ob, 0.50, 0, F, S)
    t.arrow(Ob, ue, lw=0.9)
    t.label(ue[0] + 0.01, ue[1] - 0.12, 0.12, 0.13, [T(I('u'))], size=6.8, align='left')
    ve = P(Ob, 0, 0.30, F, S)
    t.arrow(Ob, ve, lw=0.9)
    t.label(ve[0] + 0.01, ve[1] - 0.02, 0.12, 0.13, [T(I('v'))], size=6.8, align='left')
    arc(Ob, 0.075, 120, 400, n=18, head=True, lw=0.7)
    t.label(Ob[0] - 0.20, Ob[1] - 0.15, 0.10, 0.13, [T(I('r'))], size=6.6)
    # 선미 드라이브 위치와 추력
    D = P(Ob, -0.27, 0, F, S)
    delta = 25.0
    tdir = dirv(psi + delta)          # 드라이브가 향하는 방향 (추력 방향)
    ndir = dirv(psi + delta + 90)     # 드라이브 축에 수직
    t.arrow(D, P(D, -0.34, 0, F, S), color=C_GRAY, lw=0.6, head=False, dash=True)
    # 선미 드라이브 다리 (조향각 δ 만큼 돌아간 상태)와 프로펠러
    Ld, wd = 0.12, 0.036

    def Q(o, a, b):  # o + a*tdir + b*ndir
        return (o[0] + a * tdir[0] + b * ndir[0], o[1] + a * tdir[1] + b * ndir[1])

    t.polygon([Q(D, 0.01, -wd / 2), Q(D, 0.01, wd / 2), Q(D, -Ld, wd * 0.35), Q(D, -Ld, -wd * 0.35)],
              '1B3A5C', name='sterndrive_leg')
    Pp = Q(D, -Ld, 0)
    t.polygon([Q(Pp, 0.010, -0.075), Q(Pp, 0.010, 0.075), Q(Pp, -0.010, 0.075), Q(Pp, -0.010, -0.075)],
              '3A3A3A', name='propeller_blades')
    t.ellipse(Pp[0], Pp[1], 0.032, 0.032, '3A3A3A', 1.0)
    # 추력 화살표 (프로펠러 뒤에서 드라이브 방향으로)
    Fs = Q(Pp, -0.30, 0)
    t.arrow(Fs, Q(Pp, -0.02, 0), color='C0392B', lw=1.3)
    t.label(Fs[0] - 0.05, Fs[1] + 0.03, 0.20, 0.14, [T(I('F'), ISUB('g'))], size=6.8, color='C0392B', align='left')
    arc(D, 0.21, psi + 180, psi + delta + 180, n=8, color='C0392B')
    dl = (D[0] + 0.29 * dirv(psi + 180 + delta / 2)[0], D[1] + 0.29 * dirv(psi + 180 + delta / 2)[1])
    t.label(dl[0] - 0.06, dl[1] - 0.07, 0.12, 0.14, [T(I('δ'))], size=6.8, color='C0392B')
    lm = P(Ob, -0.135, 0.17, F, S)
    t.label(lm[0] - 0.08, lm[1] - 0.07, 0.16, 0.14, [T(I('l'), ISUB('d'))], size=6.4)

    out = os.path.join(OUTDIR, 'fig3_frames.pptx')
    _save(prs, out)
    return out


if __name__ == '__main__':
    print(make_fig1())
    print(make_fig2())
    print(make_fig3())
