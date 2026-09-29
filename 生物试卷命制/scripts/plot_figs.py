# -*- coding: utf-8 -*-
"""绘制八年级上册第六章单元测试卷所需示意图（简体黑白线条）。"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Ellipse, Arc, Rectangle, FancyBboxPatch, Polygon, PathPatch
from matplotlib.path import Path
import numpy as np
import os

plt.rcParams["font.sans-serif"] = ["Noto Sans CJK SC", "Noto Sans SC"]
plt.rcParams["axes.unicode_minus"] = False

OUT = "/sandbox/workspace/output/biovol6/"
os.makedirs(OUT, exist_ok=True)


def fig_eye():
    """眼球结构示意图（横向剖面）。"""
    fig, ax = plt.subplots(figsize=(5.4, 4.0), dpi=200)
    ax.set_xlim(-2.2, 2.0); ax.set_ylim(-1.6, 1.9)
    ax.set_aspect('equal'); ax.axis('off')
    # 巩膜（外层）
    ax.add_patch(Circle((0, 0), 1.0, fill=False, lw=2.0, ec='black'))
    # 视网膜（内层，虚线）
    ax.add_patch(Circle((0, 0), 0.9, fill=False, lw=1.1, ec='0.35', ls=(0, (4, 3))))
    # 角膜（前方凸弧，位于眼球右侧）
    ax.add_patch(Arc((0.35, 0), 1.7, 2.0, angle=0, theta1=-52, theta2=52, lw=2.2, ec='black'))
    # 虹膜（晶状体前方上下两片）
    ax.add_patch(Arc((0.60, 0), 0.62, 0.95, angle=0, theta1=40, theta2=78, lw=2.0, ec='black'))
    ax.add_patch(Arc((0.60, 0), 0.62, 0.95, angle=0, theta1=-78, theta2=-40, lw=2.0, ec='black'))
    # 晶状体
    ax.add_patch(Ellipse((0.46, 0), 0.26, 0.78, fill=False, lw=2.0, ec='black'))
    # 玻璃体文字
    ax.text(-0.30, 0, "玻璃体", ha='center', va='center', fontsize=11)
    # 视神经
    ax.add_patch(Rectangle((-2.1, -0.13), 1.12, 0.26, fill=False, lw=1.5, ec='black'))
    # 标注
    ax.annotate("角膜", xy=(1.15, 0.52), xytext=(1.05, 1.5), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("瞳孔", xy=(0.62, 0.0), xytext=(0.30, 1.15), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("晶状体", xy=(0.46, 0.40), xytext=(-0.15, 1.5), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("视网膜", xy=(0.30, -0.86), xytext=(0.35, -1.45), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("视神经", xy=(-1.5, 0.14), xytext=(-1.55, 1.0), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    plt.tight_layout()
    plt.savefig(OUT + "fig_eye.png", bbox_inches='tight', facecolor='white')
    plt.close()


def fig_ear():
    """耳的结构示意图。"""
    fig, ax = plt.subplots(figsize=(5.6, 3.4), dpi=200)
    ax.set_xlim(-0.8, 6.6); ax.set_ylim(-2.2, 2.2)
    ax.set_aspect('equal'); ax.axis('off')
    # 耳郭
    ax.add_patch(Arc((1.0, 0), 2.6, 3.4, angle=0, theta1=-68, theta2=112, lw=2.2, ec='black'))
    # 外耳道
    ax.plot([1.7, 3.3], [0.4, 0.4], lw=1.6, color='black')
    ax.plot([1.7, 3.3], [-0.4, -0.4], lw=1.6, color='black')
    # 鼓膜
    ax.add_patch(Ellipse((3.4, 0), 0.12, 1.15, fill=False, lw=2.0, ec='black'))
    # 听小骨（两枚简化）
    ax.add_patch(Rectangle((3.65, -0.18), 0.28, 0.36, fill=False, lw=1.6, ec='black'))
    ax.add_patch(Polygon([[3.98, -0.25], [4.30, -0.02], [3.98, 0.25]], closed=True,
                         fill=False, lw=1.6, ec='black'))
    # 耳蜗（螺旋）
    t = np.linspace(0, 3.1 * np.pi, 240)
    r = np.linspace(0.12, 0.78, 240)
    ax.plot(4.85 + r * np.cos(t), -0.25 + r * np.sin(t), lw=1.8, color='black')
    # 听神经
    ax.plot([5.30, 6.35], [-0.90, -1.55], lw=1.8, color='black')
    # 标注
    ax.annotate("耳郭", xy=(0.15, 1.35), xytext=(0.10, 1.95), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("外耳道", xy=(2.5, 0.40), xytext=(2.4, 1.4), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("鼓膜", xy=(3.40, 0.58), xytext=(3.65, 1.75), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("听小骨", xy=(4.10, 0.28), xytext=(4.30, 1.75), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("耳蜗", xy=(5.05, 0.45), xytext=(4.95, 1.60), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("听神经", xy=(5.95, -1.30), xytext=(6.05, -1.95), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    plt.tight_layout()
    plt.savefig(OUT + "fig_ear.png", bbox_inches='tight', facecolor='white')
    plt.close()


def fig_reflex():
    """反射弧结构模式图。"""
    fig, ax = plt.subplots(figsize=(6.6, 1.7), dpi=200)
    ax.set_xlim(0, 13.2); ax.set_ylim(0, 3)
    ax.axis('off')
    labels = ["感受器", "传入神经", "神经中枢", "传出神经", "效应器"]
    xs = [0.4, 2.95, 5.5, 8.05, 10.6]
    for x, lab in zip(xs, labels):
        ax.add_patch(FancyBboxPatch((x, 0.95), 2.05, 1.1,
                     boxstyle="round,pad=0.05", fill=False, lw=1.6, ec='black'))
        ax.text(x + 1.02, 1.5, lab, ha='center', va='center', fontsize=12)
    for i in range(4):
        ax.annotate("", xy=(xs[i + 1] - 0.03, 1.5), xytext=(xs[i] + 2.05 + 0.03, 1.5),
                    arrowprops=dict(arrowstyle='-|>', lw=1.8))
    plt.tight_layout()
    plt.savefig(OUT + "fig_reflex.png", bbox_inches='tight', facecolor='white')
    plt.close()


def fig_joint():
    """关节结构示意图。"""
    fig, ax = plt.subplots(figsize=(4.2, 3.8), dpi=200)
    ax.set_xlim(0, 6.4); ax.set_ylim(0, 6.4)
    ax.set_aspect('equal'); ax.axis('off')
    # 上骨（下方为关节头，圆头）
    ax.add_patch(Polygon([[2.6, 4.0], [2.6, 6.0], [3.8, 6.0], [3.8, 4.0]], closed=True,
                         fill=False, lw=1.8, ec='black'))
    ax.add_patch(Arc((3.2, 4.0), 1.2, 1.2, angle=0, theta1=180, theta2=360, lw=1.8, ec='black'))
    # 关节软骨（关节头表面）
    ax.add_patch(Arc((3.2, 4.0), 1.34, 1.34, angle=0, theta1=185, theta2=355, lw=2.4, ec='0.35'))
    # 下骨（上方为关节窝，凹面）
    ax.add_patch(Polygon([[2.6, 0.6], [2.6, 3.4], [3.8, 3.4], [3.8, 0.6]], closed=True,
                         fill=False, lw=1.8, ec='black'))
    ax.add_patch(Arc((3.2, 3.4), 1.2, 1.2, angle=0, theta1=185, theta2=355, lw=1.8, ec='black'))
    # 关节囊（包裹）
    ax.add_patch(Arc((3.2, 3.7), 2.0, 2.2, angle=0, theta1=150, theta2=30, lw=1.8, ec='black'))
    ax.add_patch(Arc((3.2, 3.7), 2.0, 2.2, angle=0, theta1=330, theta2=210, lw=1.8, ec='black'))
    # 标注
    ax.annotate("关节头", xy=(2.75, 3.95), xytext=(0.85, 4.7), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("关节窝", xy=(3.75, 3.45), xytext=(5.7, 3.0), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("关节软骨", xy=(3.2, 4.72), xytext=(5.6, 5.2), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("关节囊", xy=(2.25, 3.7), xytext=(0.8, 2.6), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    ax.annotate("关节腔", xy=(3.2, 3.5), xytext=(1.0, 1.6), fontsize=11, ha='center',
                arrowprops=dict(arrowstyle='-', lw=1.0))
    plt.tight_layout()
    plt.savefig(OUT + "fig_joint.png", bbox_inches='tight', facecolor='white')
    plt.close()


if __name__ == "__main__":
    fig_eye(); fig_ear(); fig_reflex(); fig_joint()
    print("figures done")
    for f in ["fig_eye.png", "fig_ear.png", "fig_reflex.png", "fig_joint.png"]:
        print(f, os.path.getsize(OUT + f), "bytes")
