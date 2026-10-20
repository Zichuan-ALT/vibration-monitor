# 主题 8.1 · 震动状态监测器

> 利莫瑞克大学 × 河南理工大学（LICHPU）　工业评估（15 ECTS）
> 智慧校园生活实验室 · 主题 8「智能设施与预测性维护」

一套面向校园旋转设备（电机 / 风机）的**低成本、免焊接、可复制**振动状态监测节点。
节点在边缘完成特征提取与状态判定，通过 MQTT 上报；服务端提供实时仪表板、
SQLite 历史存储与分级维修建议。

---

## 当前进度

| 环节 | 状态 | 说明 |
|---|---|---|
| 提案与需求规格 | ✅ 完成 | 11 项 FR / 9 项 NFR / 13 条 AC |
| 实施方案 | ✅ 完成 | 13 个施工步骤 + 附录 A 环境故障速查 |
| 受控文档台账 | ✅ 完成 | 9 个工作表，成本 ¥119.50 / 结余 ¥30.50 |
| ESP32 工具链 | ✅ 打通 | Arduino IDE + esp32 core 2.0.17 |
| Wi-Fi 连接 | ✅ 打通 | 2.4 GHz 热点，RSSI −34 dBm |
| MQTT 上报 | ✅ 打通 | ESP32 → aedes broker → Node-RED |
| 仪表板（实时 / 历史） | ✅ 打通 | 2 张仪表盘 + 2 条曲线 |
| SQLite 历史存储 | ✅ 打通 | readings 表，每 3 s 一条 |
| 报警状态机（3/5 帧迟滞） | ✅ 打通 | normal / warn / alarm + 维修建议 |
| **真实传感器接入** | ⏳ 待硬件 | 等杜邦线与已焊针 MPU6050 |
| 采样 + FFT 特征值 | ⏳ 待硬件 | 当前用模拟工况数据验证链路 |

> 当前固件以**模拟工况**（90 s 周期、RMS 0.010→0.070）验证整条链路与报警状态机，
> 传感器到位后仅替换数据源，服务端无需改动。

---

## 目录结构

```
.
├── 8.1_振动状态监测器_项目提案与需求规格说明.docx / .pdf   # 交付物 A
├── 8.1_振动状态监测器_实施方案.docx / .pdf                # 交付物 B
├── 8.1_受控文档台账.xlsx                                 # 交付物 C
├── build_proposal2.py                                   # 提案生成脚本
├── build_simple3.py                                     # 实施方案生成脚本
├── build_workbook3.py                                   # 台账生成脚本
├── figures/                                             # 插图与绘图脚本
├── firmware/                                            # ESP32 固件（按步骤编号）
├── nodered/                                             # Node-RED 流程与依赖
├── evidence/                                            # 运行证据截图
└── archive/                                             # 早期草稿（保留设计演进记录）
```

---

## 快速复现

### 1. 硬件

| 器件 | 型号 | 数量 |
|---|---|---|
| 主控 | ESP32-WROOM-32（CH340C，已焊针 30Pin，Type-C） | 2（1 备） |
| 振动 + 板载温度 | MPU-6050 三轴加速度模块（GY-521，已焊针） | 1 |
| 表面温度 | DS18B20 带防水探头模块 | 1 |
| 存储 | Micro SD 卡模块（SPI，已焊针）+ 32 GB Class10 A1 | 各 1 |
| 指示 | 有源蜂鸣器模块（3.3–5 V，低电平触发） | 1 |
| 演示设备 | 静音风扇 5 V 9225（1700 RPM） | 1 |
| 其他 | 杜邦线（学校提供）、面包板、USB 充电器 ×2 | — |

总成本 **¥119.50**，预算上限 ¥150，结余 ¥30.50。

### 2. 开发环境

Arduino IDE + **esp32 core 2.0.17**（不要用 3.x，体积 1.9 GB 且 API 有变）。
国内网络下的安装故障与对策见实施方案 **附录 A**。

必需的库：

```
PubSubClient 2.8.0        MQTT 客户端
arduinoFFT                FFT
OneWire + DallasTemperature   DS18B20
MPU6050 (Electronic Cats)     三轴加速度
```

### 3. 服务端

```powershell
npm config set registry https://registry.npmmirror.com
npm install -g node-red
cd nodered
npm install
node-red -u .
```

- 编辑器：<http://localhost:1880>
- 仪表板：<http://localhost:1880/ui>
- MQTT broker：`1883`

### 4. 固件烧录

按 `firmware/` 编号顺序：

| 目录 | 用途 |
|---|---|
| `01_blink` | 最小系统验证（板载 LED） |
| `02_wifi_scan` | 2.4 GHz 网络扫描 |
| `03_wifi_connect` | Wi-Fi 连接与 IP 获取 |
| `04_mqtt_publish` | MQTT 上报（含模拟工况与报警演示） |

烧录前把 `04_mqtt_publish` 里的 `SSID` / `PASS` / `BROKER` 改成你的实际值
（`BROKER` 填运行 Node-RED 那台机器的局域网 IP）。

---

## 通信协议

主题：`smartcampus/vib/{node_id}/state`（QoS 1）

```json
{
  "node":  "esp32-test",
  "seq":   42,
  "rms":   0.0158,
  "temp":  26.7,
  "state": "normal"
}
```

服务端另订阅 `smartcampus/vib/+/event` 接收事件消息。

---

## 报警判据

| 等级 | 条件（演示用固定阈值） | 维修建议 |
|---|---|---|
| normal | RMS < 0.030 g | 运行正常，继续保持监测 |
| warn | 0.030 ≤ RMS < 0.060 g | 检查地脚螺栓与紧固件 |
| alarm | RMS ≥ 0.060 g | 停机检查动平衡与轴承 |

正式版用**基线自学习**（开机 60 s 学习得到 μ、σ）动态设定 μ+3σ / μ+6σ，
并采用**迟滞**：升级需连续 3 帧确认、降级需连续 5 帧确认，抑制阈值附近抖动。

---

## 文档重新生成

三份文档均由脚本生成，改动后重新运行即可（需要 python-docx / openpyxl）：

```powershell
python build_proposal2.py      # 提案与需求规格说明
python build_simple3.py        # 实施方案
python build_workbook3.py      # 受控文档台账
python figures/draw_figures2.py   # 插图
```

---

## 说明

- 本仓库的 `firmware/` 是开发过程快照；实际调试在 Arduino IDE 的项目目录中进行。
- `archive/` 保留早期草稿，用于记录设计决策的演进（评估要求「受控文档」）。
- 与本项目无关的 `LH4018 Lab05` 课程材料已通过 `.gitignore` 排除。
