# -*- coding: utf-8 -*-
"""
數學科「座標先行 (Coordinate-First)」300 DPI 高清繪圖引擎 (math_geometry_plotter.py)
嚴格遵守：
1. 平面幾何圖強制等比例鎖定 (ax.set_aspect('equal') + ax.axis('off'))
2. 頂點字母外推防壓線演算法 (Smart Vertex Offset)
3. 必備幾何標記：直角記號、角度弧線、等長標記、斜線陰影
4. 十字直角坐標系規範 (原點 O、軸箭頭、去四邊封閉方框)
5. +8 Pt 特大清晰字級 (16.5~21 Pt)，黑白灰階高對比印刷優化
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DFKai-SB', 'Times New Roman']
plt.rcParams['axes.unicode_minus'] = False

def create_math_figure(figsize=(4.8, 4.2), dpi=300):
    """建立數學專用白底高清 300 DPI 畫布"""
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    return fig, ax

def draw_right_angle(ax, vertex, p1, p2, size=0.35, lw=1.8):
    """在 vertex 處繪製指向 p1 與 p2 的標準直角記號 (小方框)"""
    v = np.array(vertex, dtype=float)
    u1 = np.array(p1, dtype=float) - v
    u2 = np.array(p2, dtype=float) - v
    u1 = u1 / np.linalg.norm(u1) * size
    u2 = u2 / np.linalg.norm(u2) * size
    c1 = v + u1
    c2 = v + u1 + u2
    c3 = v + u2
    ax.plot([c1[0], c2[0], c3[0]], [c1[1], c2[1], c3[1]], color='black', lw=lw)

def draw_angle_arc(ax, vertex, p1, p2, radius=0.6, label=None, fontsize=17):
    """在 vertex 處繪製從 p1 到 p2 (逆時針) 的角度圓弧與度數標籤"""
    v = np.array(vertex, dtype=float)
    d1 = np.array(p1, dtype=float) - v
    d2 = np.array(p2, dtype=float) - v
    ang1 = np.degrees(np.arctan2(d1[1], d1[0])) % 360
    ang2 = np.degrees(np.arctan2(d2[1], d2[0])) % 360
    if (ang2 - ang1) % 360 > 180:
        ang1, ang2 = ang2, ang1
    arc = patches.Arc(v, 2*radius, 2*radius, angle=0, theta1=ang1, theta2=ang2, color='black', lw=1.8)
    ax.add_patch(arc)
    if label:
        mid_ang = np.radians(ang1 + ((ang2 - ang1) % 360) / 2.0)
        lx = v[0] + (radius * 1.48) * np.cos(mid_ang)
        ly = v[1] + (radius * 1.48) * np.sin(mid_ang)
        ax.text(lx, ly, label, fontsize=fontsize, ha='center', va='center', fontweight='bold')

def draw_tick_marks(ax, p1, p2, num_ticks=1, tick_len=0.22, spacing=0.12, lw=1.8):
    """在線段 p1-p2 中點繪製等長記號 (單槓/雙槓/三槓)"""
    p1, p2 = np.array(p1, float), np.array(p2, float)
    mid = (p1 + p2) / 2.0
    diff = p2 - p1
    dist = np.linalg.norm(diff)
    if dist < 1e-6:
        return
    tangent = diff / dist
    normal = np.array([-tangent[1], tangent[0]])
    offsets = np.linspace(-(num_ticks-1)*spacing/2, (num_ticks-1)*spacing/2, num_ticks)
    for off in offsets:
        center = mid + off * tangent
        a = center - (tick_len / 2) * normal
        b = center + (tick_len / 2) * normal
        ax.plot([a[0], b[0]], [a[1], b[1]], color='black', lw=lw)

def label_vertex(ax, pt, name, center_ref=(0, 0), offset=0.42, fontsize=19):
    """沿著遠離圖形中心 center_ref 的方向自動外推標註頂點字母，絕不壓線"""
    pt = np.array(pt, float)
    c = np.array(center_ref, float)
    vec = pt - c
    norm = np.linalg.norm(vec)
    unit = vec / norm if norm > 1e-6 else np.array([0.0, 1.0])
    pos = pt + unit * offset
    ax.text(pos[0], pos[1], f"${name}$", fontsize=fontsize, ha='center', va='center', fontweight='bold')

def setup_cartesian_axes(ax, xlim=(-5, 5), ylim=(-5, 5)):
    """設定標準中學數學十字直角座標系 (含原點 O 與 x, y 軸箭頭)"""
    ax.set_xlim(xlim)
    ax.set_ylim(ylim)
    ax.set_aspect('equal')
    ax.spines['left'].set_position('zero')
    ax.spines['bottom'].set_position('zero')
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['left'].set_linewidth(1.8)
    ax.spines['bottom'].set_linewidth(1.8)
    ax.plot(xlim[1], 0, ">k", markersize=8, clip_on=False)
    ax.plot(0, ylim[1], "^k", markersize=8, clip_on=False)
    ax.text(xlim[1] + 0.3, -0.1, "$x$", fontsize=19, fontweight='bold', va='center')
    ax.text(0.15, ylim[1] + 0.3, "$y$", fontsize=19, fontweight='bold', ha='center')
    ax.text(-0.35, -0.4, "$O$", fontsize=17, fontweight='bold')

def draw_shaded_polygon(ax, vertices, hatch='///', facecolor='#E0E0E0', edgecolor='black', lw=1.5):
    """繪製面積求值之陰影多邊形 (淺灰底色搭配黑白斜線)"""
    poly = patches.Polygon(vertices, closed=True, facecolor=facecolor, hatch=hatch, edgecolor=edgecolor, lw=lw)
    ax.add_patch(poly)
