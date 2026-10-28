# 证据文件说明

本目录存放报告的实物证据与界面截图。

## 实物照片

| 文件 | 内容 | 用途 |
|---|---|---|
| `setup_overview.jpg` | 整体装置：风扇 + MPU6050 + DS18B20 + ESP32 + 面包板 + SD 模块 | 报告"系统实现"章节，说明整体形态 |
| `mpu6050_mount.jpg` | 振动传感器安装特写：模块用胶带固定于风扇外框，电源 LED 点亮 | 说明传感器安装方式与刚性固定要求 |
| `ds18b20_probe.jpg` | 温度探头位置：探头置于进风面前方、电机轮毂边缘 | 说明温度测点位置（温度读数与位置强相关） |
| `wiring_breadboard.jpg` | 面包板接线全貌：ESP32、SD 模块、各传感器连线 | 供复现参考 |
| `blockage_test.jpg` | 风道堵塞实验现场：纸板遮挡风扇进风面 | 风道堵塞工况的实验证据 |
| `device_with_dashboard.jpg` | 实物与仪表板同框：风扇运转，笔记本显示实时数据 | 证明仪表板数据来自眼前设备 |

## 界面截图

| 文件 | 内容 |
|---|---|
| `dashboard_live_normal.png` | 正常工况仪表板（单通道时期） |
| `dashboard_live_alarm.png` | 报警工况仪表板，含趋势阶跃 |
| `dashboard_dual_sensor.png` | 双通道仪表板（振动 + 表面温度） |
| `dashboard_alarm.png` | 模拟数据时期的报警界面（早期） |
| `dashboard_live_v1.png` / `dashboard_live_v2.png` | 早期版本（归档） |

## 数据图表

频谱对比、历史曲线等由脚本生成，存放在 `experiments/data/`：

| 文件 | 内容 | 生成脚本 |
|---|---|---|
| `spectra_compare.png` | 三工况频谱对比 | `tools/plot_spectra.py` |
| `spectrum_工况*.png` | 各工况单张频谱 | 同上 |
| `history_from_db.png` | SQLite 历史曲线 | `tools/plot_history.py` |

## 说明

- 照片为现场手机拍摄，未做后期处理，保留原始状态
- `setup_overview.jpg` 与 `device_with_dashboard.jpg` 背景较杂乱、
  屏幕曝光偏亮，但不影响内容辨识，故保留
