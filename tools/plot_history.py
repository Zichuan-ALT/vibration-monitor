# -*- coding: utf-8 -*-
"""从 SQLite 导出历史数据并绘图（报告证据用）。

用法：
    python tools/plot_history.py

输出：
    experiments/data/history_from_db.png
"""

import os
import sqlite3
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

DB = os.path.expanduser(r"~\.node-red\vibration.db")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "experiments", "data", "history_from_db.png")


def main():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute("SELECT ts, rms, state FROM readings ORDER BY id")
    rows = cur.fetchall()
    conn.close()

    if not rows:
        print("数据库里没有数据")
        return

    # ts 是 UTC ISO 字符串，转换为本地时间
    times, values, states = [], [], []
    for ts, rms, state in rows:
        s = ts.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s).astimezone().replace(tzinfo=None)
        times.append(dt)
        values.append(rms)
        states.append(state)

    print("共 %d 条，时间范围 %s ~ %s" % (
        len(rows), times[0].strftime("%H:%M:%S"), times[-1].strftime("%H:%M:%S")))

    fig, ax = plt.subplots(figsize=(13, 5))
    ax.plot(times, values, lw=1.2, color="#1f77b4", label="带内 RMS")

    # 标注报警区间
    alarm_t = [t for t, s in zip(times, states) if s == "alarm"]
    alarm_v = [v for v, s in zip(values, states) if s == "alarm"]
    if alarm_t:
        ax.scatter(alarm_t, alarm_v, s=10, color="#d62728", zorder=3,
                   label="报警状态 (%d 点)" % len(alarm_t))

    ax.set_xlabel("时间")
    ax.set_ylabel("带内 RMS (g)")
    ax.set_title("振动 RMS 历史曲线（数据源：SQLite readings 表）")
    ax.grid(alpha=0.3)
    ax.legend(loc="upper left")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S", tz=None))
    fig.autofmt_xdate()
    fig.tight_layout()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    fig.savefig(OUT, dpi=150)
    plt.close(fig)
    print("已输出:", OUT)

    # 简要统计
    normal = [v for v, s in zip(values, states) if s == "normal"]
    if normal:
        print("normal  RMS: %.4f ~ %.4f g (n=%d)" %
              (min(normal), max(normal), len(normal)))
    if alarm_v:
        print("alarm   RMS: %.4f ~ %.4f g (n=%d)" %
              (min(alarm_v), max(alarm_v), len(alarm_v)))


if __name__ == "__main__":
    main()
