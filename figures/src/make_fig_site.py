"""Fig. 9 (experimental site) with the inset taken from the overhead drone video of 2026-11-12.

Main photo: photos/site_base.png, the contrast enhanced crop of test_site.png that the earlier
fig_site.pdf embedded (extracted from it unchanged), with the zoom box and labels at the same
positions as before.

Inset: photos/drone_run27_v156.1_crop.png, a 900 x 600 crop (origin u = 582, v = 241) of
EXP_video/output_1080p/2026-11-12_1112_DJI_0384-0389_1080p.mp4 at 156.1 s, during run 27
(DC-RIS), when the ASV was 1.54 m from the setpoint (median over the drone segment 1.62 m).
The keep in circle and the setpoint are drawn to scale. Georeferencing (scratchpad fig9/georef.py):
on ten frames over the 309 s segment the white hull was segmented, scale = hull length / 8.29 m
(46.62 px/m, the same scale gives a 2.5 m beam), rotation = image bow angle - EKF heading
(-99.0 deg), and the hull centre was matched to the EKF position. The setpoint pixel implied by
the ten frames scatters by about 0.17 m, so the drone held its position.

    python make_fig_site.py      ->  ../fig_site_drone.pdf  and  render_fig_site_drone.png
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import matplotlib.patheffects as pe
from matplotlib.patches import Rectangle, Circle

HERE = os.path.dirname(os.path.abspath(__file__))
PH = os.path.join(HERE, "photos")
W, H = 252.0, 139.667                         # pt, single column (same as the old fig_site.pdf)
YEL = "#FFD34D"

# georeferencing result, in pixels of the crop
SCALE = 46.62                                 # px per m
SP_U, SP_V = 1031.9 - 582, 540.8 - 241        # setpoint
R_KEEP = 4.0

plt.rcParams.update({"font.family": "serif", "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
                     "font.size": 7, "mathtext.fontset": "stix"})
OUT = [pe.withStroke(linewidth=1.6, foreground="black")]

fig = plt.figure(figsize=(W / 72, H / 72))
ax = fig.add_axes([0, 0, 1, 1])
ax.imshow(mpimg.imread(os.path.join(PH, "site_base.png")), extent=[0, W, H, 0], interpolation="none")
ax.set_xlim(0, W); ax.set_ylim(H, 0); ax.axis("off")

# zoom box on the ASV in the main photo (same place as before)
bx0, bx1, by0, by1 = 70.5, 123.8, 70.2, 96.8
ax.add_patch(Rectangle((bx0, by0), bx1 - bx0, by1 - by0, fill=False, ec="white", lw=0.6))
ax.text((bx0 + bx1) / 2, by0 - 3.2, "ASV", color="white", ha="center", va="bottom", fontsize=7.5, path_effects=OUT)

# inset
ix0, ix1, iy1 = 147.0, 247.0, 134.7
HU, HV = 340.0, 215.0                         # half extent of the shown region, px
iy0 = iy1 - (ix1 - ix0) * HV / HU
for (xa, ya), (xb, yb) in (((bx1, by0), (ix0, iy0)), ((bx1, by1), (ix0, iy1))):
    ax.plot([xa, xb], [ya, yb], color="white", lw=0.6)
axi = fig.add_axes([ix0 / W, 1 - iy1 / H, (ix1 - ix0) / W, (iy1 - iy0) / H])
img = mpimg.imread(os.path.join(PH, "drone_run27_v156.1_crop.png"))[..., :3].astype(float)
# the drone footage is flat, so stretch all channels together (hue kept) between the 1 and 99.8 percentiles
reg = img[int(SP_V - HV):int(SP_V + HV), int(SP_U - HU):int(SP_U + HU)]
lo, hi = np.percentile(reg, 1.0), np.percentile(reg, 99.8)
img = np.clip(0.15 + 0.85 * (img - lo) / (hi - lo), 0, 1)
axi.imshow(img, interpolation="none")
axi.set_xlim(SP_U - HU, SP_U + HU); axi.set_ylim(SP_V + HV, SP_V - HV)
axi.set_xticks([]); axi.set_yticks([])
for s in axi.spines.values():
    s.set_color("white"); s.set_linewidth(0.8)

rk = R_KEEP * SCALE
axi.add_patch(Circle((SP_U, SP_V), rk, fc=YEL, alpha=0.12, ec="none"))
axi.add_patch(Circle((SP_U, SP_V), rk, fill=False, ec=YEL, lw=1.0))
a = np.radians(38.0)                          # radius drawn toward the open water, down right
ue, ve = SP_U + rk * np.cos(a), SP_V + rk * np.sin(a)
axi.plot([SP_U, ue], [SP_V, ve], color=YEL, lw=0.8)
axi.text(SP_U + 0.62 * rk * np.cos(a) + 10, SP_V + 0.62 * rk * np.sin(a) + 16, r"$R$", color="white",
         ha="left", va="top", fontsize=7.5, path_effects=OUT)
axi.plot([SP_U], [SP_V], "o", ms=2.6, mfc=YEL, mec="black", mew=0.4)
lu, lv = SP_U - 175, SP_V + 140
axi.plot([SP_U, lu + 40], [SP_V, lv - 22], color="white", lw=0.5)
axi.text(lu, lv, "Setpoint", color="white", ha="center", va="top", fontsize=7, path_effects=OUT)
axi.text(SP_U - rk * 0.95, SP_V - rk * 0.80, "Keep in circle", color="white", ha="center", va="bottom",
         fontsize=7, path_effects=OUT)

fig.savefig(os.path.join(HERE, "..", "fig_site_drone.pdf"))
fig.savefig(os.path.join(HERE, "render_fig_site_drone.png"), dpi=300)
print("wrote fig_site_drone.pdf")
