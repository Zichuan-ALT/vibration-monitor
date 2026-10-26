# -*- coding: utf-8 -*-
"""把 experiments/data 下的频谱 CSV 画成图。

用法：
    python tools/plot_spectra.py

读取 experiments/data/spectrum_*.csv（文件首行可带 "# label: 工况说明"），
输出 experiments/data/spectra_compare.png 与 spectra_detail_<label>.png
"""

import glob
import os
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "experiments", "data")
BAND = (20.0, 45.0)
ANALYSIS = (10.0, 260.0)


def load(path):
    """读取一个频谱 CSV，返回 (freq, amp, meta)。"""
    meta = {"label": os.path.splitext(os.path.basename(path))[0]}
    freqs, amps = [], []
    with open(path, encoding="utf-8-sig", errors="ignore") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            if line.startswith("#"):
                m = re.match(r"#\s*label\s*[:：]\s*(.+)", line)
                if m:
                    meta["label"] = m.group(1).strip()
                m2 = re.match(r"#BEGIN,(\d+),([\d.]+),([\d.]+),([\d.]+)", line)
                if m2:
                    meta["axis"] = ["ax", "ay", "az"][int(m2.group(1))]
                    meta["rms"] = float(m2.group(2))
                    meta["peak_f"] = float(m2.group(3))
                    meta["peak_a"] = float(m2.group(4))
                continue
            parts = line.split(",")
            if len(parts) != 2:
                continue
            try:
                freqs.append(float(parts[0]))
                amps.append(float(parts[1]))
            except ValueError:
                continue
    return np.array(freqs), np.array(amps), meta


def main():
    files = sorted(glob.glob(os.path.join(BASE, "spectrum_*.csv")))
    if not files:
        print("没有找到频谱文件：", os.path.join(BASE, "spectrum_*.csv"))
        return

    data = []
    for p in files:
        f, a, m = load(p)
        if len(f) > 100:
            data.append((f, a, m))
            print("载入 %-40s  %d 点  %s" %
                  (os.path.basename(p), len(f), m.get("label", "")))

    if not data:
        print("没有有效的频谱数据")
        return

    # ---------- 图 1：多工况对比（全带 + 放大） ----------
    fig, axes = plt.subplots(1, 2, figsize=(15, 5.5))

    for f, a, m in data:
        label = m.get("label", "")
        if "rms" in m:
            label = "%s  (RMS %.4f g)" % (label, m["rms"])
        axes[0].plot(f, a, lw=0.8, label=label)
        axes[1].plot(f, a, lw=0.9, label=label)

    for ax, xlim, title in (
        (axes[0], (0, 500), "全带 0–500 Hz"),
        (axes[1], (0, 120), "低频段放大 0–120 Hz"),
    ):
        ax.axvspan(*BAND, color="#ff7f0e", alpha=0.12, label="基频带 20–45 Hz")
        ax.axvline(ANALYSIS[0], color="#2ca02c", lw=1.2, ls="--", label="分析带宽 10-260 Hz")
        ax.axvline(ANALYSIS[1], color="#2ca02c", lw=1.2, ls="--")
        ax.set_xlim(*xlim)
        ax.set_xlabel("频率 (Hz)")
        ax.set_ylabel("幅值 (g)")
        ax.set_title(title)
        ax.grid(alpha=0.3)
        ax.legend(fontsize=8, loc="upper right")

    fig.suptitle("振动加速度幅值谱（1024 点 @ 1 kSPS，Hann 窗）", fontsize=13)
    fig.tight_layout()
    out1 = os.path.join(BASE, "spectra_compare.png")
    fig.savefig(out1, dpi=150)
    plt.close(fig)
    print("已输出:", out1)

    # ---------- 图 2：每个工况单独一张 ----------
    for f, a, m in data:
        label = m.get("label", "spectrum")
        fig, ax = plt.subplots(figsize=(11, 4.5))
        ax.plot(f, a, lw=0.9, color="#1f77b4")
        ax.axvspan(*BAND, color="#ff7f0e", alpha=0.12, label="基频带 20–45 Hz")
        ax.set_xlim(0, 500)
        ax.set_xlabel("频率 (Hz)")
        ax.set_ylabel("幅值 (g)")
        title = label
        if "rms" in m:
            title += "   带内 RMS %.4f g" % m["rms"]
        if "peak_f" in m:
            title += "   峰值 %.4f g @ %.1f Hz" % (m["peak_a"], m["peak_f"])
            ax.annotate("%.1f Hz" % m["peak_f"],
                        xy=(m["peak_f"], m["peak_a"]),
                        xytext=(m["peak_f"] + 15, m["peak_a"]),
                        arrowprops=dict(arrowstyle="->", color="#d62728"),
                        color="#d62728", fontsize=9)
        ax.set_title(title)
        ax.grid(alpha=0.3)
        ax.legend(fontsize=8)
        fig.tight_layout()
        safe = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", label)[:40]
        out = os.path.join(BASE, "spectrum_%s.png" % safe)
        fig.savefig(out, dpi=150)
        plt.close(fig)
        print("已输出:", out)


if __name__ == "__main__":
    main()
