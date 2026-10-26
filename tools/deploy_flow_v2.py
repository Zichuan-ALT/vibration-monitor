# -*- coding: utf-8 -*-
"""更新流程：改用 ESP32 上报的 state 与阈值；调整仪表盘量程到 0~4 g。"""
import json
import urllib.request

TAB = "tab_vib"
AEDES = "aedes_broker_1"
BROKER_CFG = "broker_conf_1"
SQLDB = "sqlite_db_1"
LOG = "C:/Users/\u5b59\u826f\u8a00/.node-red/mqtt_received.log"
DB = "C:/Users/\u5b59\u826f\u8a00/.node-red/vibration.db"

FUNC_SPLIT = (
    "var d = msg.payload;\n"
    "if (typeof d === 'string') { try { d = JSON.parse(d); } catch (e) { return null; } }\n"
    "if (!d || d.rms === undefined) { return null; }\n"
    "return [\n"
    "  { payload: d.rms },\n"
    "  { payload: d.tchip },\n"
    "  { payload: d.state },\n"
    "  { payload: d.rms, topic: 'RMS' }\n"
    "];"
)

FUNC_DB = (
    "var d = msg.payload;\n"
    "if (typeof d === 'string') { try { d = JSON.parse(d); } catch (e) { return null; } }\n"
    "if (!d || d.rms === undefined) { return null; }\n"
    "msg.params = [ new Date().toISOString(), d.node || 'unknown', d.rms, d.tchip, d.state ];\n"
    "return msg;"
)

# 判定交给 ESP32（边缘计算），服务端只做映射与展示
FUNC_ALARM = (
    "var d = msg.payload;\n"
    "if (typeof d === 'string') { try { d = JSON.parse(d); } catch (e) { return null; } }\n"
    "if (!d) { return null; }\n"
    "\n"
    "var s = d.state || 'normal';\n"
    "var LEVEL = { normal: '正常', warn: '注意', alarm: '报警' };\n"
    "var ADVICE = {\n"
    "  normal: '运行正常，继续保持监测',\n"
    "  warn:   '振动偏高：建议检查地脚螺栓与紧固件',\n"
    "  alarm:  '振动异常：建议停机检查动平衡与轴承'\n"
    "};\n"
    "\n"
    "return [\n"
    "  { payload: LEVEL[s] || s },\n"
    "  { payload: ADVICE[s] || '' },\n"
    "  { payload: s }\n"
    "];"
)

FUNC_HIST = "msg.params = [];\nreturn msg;"

FUNC_CHART = (
    "var rows = msg.payload;\n"
    "if (!Array.isArray(rows)) { return null; }\n"
    "if (rows.length === 0) { return { payload: [] }; }\n"
    "rows = rows.slice().reverse();\n"
    "var data = rows.map(function (r) {\n"
    "  return [ Date.parse(r.ts), r.rms ];\n"
    "});\n"
    "// node-red-dashboard 折线图要求 [{series, data, labels}] 格式，\n"
    "// data 为 [[x, y], ...]；仅发 {x,y} 会触发 Bad data inject\n"
    "return { payload: [{ series: ['振动 RMS (g)'], data: data }] };\n"
)

flows = [
    {"id": TAB, "type": "tab", "label": "振动监测", "disabled": False,
     "info": "主题 8.1 振动状态监测器", "env": []},

    {"id": "tab_live", "type": "ui_tab", "name": "实时状态", "icon": "dashboard",
     "order": 1, "disabled": False, "hidden": False},
    {"id": "tab_hist", "type": "ui_tab", "name": "历史查询", "icon": "history",
     "order": 2, "disabled": False, "hidden": False},
    {"id": "grp_real", "type": "ui_group", "name": "实时状态", "tab": "tab_live",
     "order": 1, "disp": True, "width": "12", "collapse": False},
    {"id": "grp_alarm", "type": "ui_group", "name": "报警与建议", "tab": "tab_live",
     "order": 2, "disp": True, "width": "12", "collapse": False},
    {"id": "grp_hist", "type": "ui_group", "name": "趋势", "tab": "tab_live",
     "order": 3, "disp": True, "width": "12", "collapse": False},
    {"id": "grp_db", "type": "ui_group", "name": "数据库历史", "tab": "tab_hist",
     "order": 1, "disp": True, "width": "12", "collapse": False},

    {"id": AEDES, "type": "aedes broker", "z": TAB,
     "name": "MQTT Broker (1883)", "mqtt_port": 1883,
     "mqtt_ws_bind": "port", "mqtt_ws_port": None, "mqtt_ws_path": "",
     "cert": "", "key": "", "ca": "", "certname": "", "keyname": "", "caname": "",
     "persistence_bind": "memory", "dburl": "", "persist_to_file": False,
     "usetls": False, "x": 160, "y": 60, "wires": []},

    {"id": BROKER_CFG, "type": "mqtt-broker", "name": "本地 aedes 1883",
     "broker": "127.0.0.1", "port": "1883", "clientid": "",
     "autoConnect": True, "usetls": False, "protocolVersion": "4",
     "keepalive": "60", "cleansession": True,
     "birthTopic": "", "birthQos": "0", "birthPayload": "", "birthMsg": {},
     "closeTopic": "", "closeQos": "0", "closePayload": "", "closeMsg": {},
     "willTopic": "", "willQos": "0", "willPayload": "", "willMsg": {}},

    {"id": "mqtt_in_1", "type": "mqtt in", "z": TAB,
     "name": "ESP32 /state", "topic": "smartcampus/vib/+/state", "qos": "1",
     "datatype": "json", "broker": BROKER_CFG,
     "nl": False, "rap": True, "rh": 0, "inputs": 0,
     "x": 150, "y": 150, "wires": [["func_1", "func_db", "func_alarm", "file_1", "debug_1"]]},

    {"id": "func_1", "type": "function", "z": TAB, "name": "拆分字段",
     "func": FUNC_SPLIT, "outputs": 4, "timeout": 0, "noerr": 0,
     "initialize": "", "finalize": "", "libs": [],
     "x": 380, "y": 40,
     "wires": [["gauge_rms"], ["gauge_temp"], ["text_state"], ["chart_rms"]]},

    # 量程改为 0~4 g（实测 RMS 从 0.05 到 3 g）
    {"id": "gauge_rms", "type": "ui_gauge", "z": TAB, "name": "振动 RMS",
     "group": "grp_real", "order": 1, "width": "6", "height": "5",
     "gtype": "gage", "title": "振动 RMS (g)　量程 0-0.5", "label": "g",
     "format": "{{value}}", "min": 0, "max": "0.5",
     "colors": ["#00b500", "#e6e600", "#ca3838"], "seg1": "", "seg2": "",
     "x": 640, "y": 10, "wires": []},

    {"id": "gauge_temp", "type": "ui_gauge", "z": TAB, "name": "芯片温度",
     "group": "grp_real", "order": 2, "width": "6", "height": "5",
     "gtype": "gage", "title": "MPU6050 芯片温度 (℃)", "label": "℃",
     "format": "{{value}}", "min": 0, "max": "60",
     "colors": ["#00b500", "#e6e600", "#ca3838"], "seg1": "", "seg2": "",
     "x": 640, "y": 70, "wires": []},

    {"id": "text_state", "type": "ui_text", "z": TAB, "name": "设备上报状态",
     "group": "grp_real", "order": 3, "width": "12", "height": "2",
     "name2": "状态", "label": "设备上报状态（ESP32 边缘判定）",
     "format": "{{msg.payload}}", "layout": "row-spread",
     "x": 640, "y": 130, "wires": []},

    {"id": "func_alarm", "type": "function", "z": TAB, "name": "状态映射（来自设备）",
     "func": FUNC_ALARM, "outputs": 3, "timeout": 0, "noerr": 0,
     "initialize": "", "finalize": "", "libs": [],
     "x": 400, "y": 160,
     "wires": [["text_alarm"], ["text_advice"], []]},

    {"id": "text_alarm", "type": "ui_text", "z": TAB, "name": "报警等级",
     "group": "grp_alarm", "order": 1, "width": "6", "height": "2",
     "name2": "等级", "label": "报警等级", "format": "{{msg.payload}}",
     "layout": "row-spread", "x": 700, "y": 150, "wires": []},

    {"id": "text_advice", "type": "ui_text", "z": TAB, "name": "维修建议",
     "group": "grp_alarm", "order": 2, "width": "12", "height": "2",
     "name2": "建议", "label": "维修建议", "format": "{{msg.payload}}",
     "layout": "row-spread", "x": 700, "y": 200, "wires": []},

    {"id": "chart_rms", "type": "ui_chart", "z": TAB, "name": "RMS 实时趋势",
     "group": "grp_hist", "order": 1, "width": "12", "height": "6",
     "label": "振动 RMS 实时趋势", "chartType": "line", "legend": False,
     "xformat": "HH:mm:ss", "interpolate": "linear", "nodata": "等待数据…",
     "ymin": "0", "ymax": "4",
     "removeOlder": "60", "removeOlderPoints": "", "removeOlderUnit": "60",
     "cutout": 0, "useOneColor": True, "useUTC": False,
     "colors": ["#1f77b4"], "outputs": 1, "x": 640, "y": 250, "wires": [[]]},

    {"id": SQLDB, "type": "sqlitedb", "db": DB, "mode": "RWC"},

    {"id": "func_db", "type": "function", "z": TAB, "name": "组装 INSERT 参数",
     "func": FUNC_DB, "outputs": 1, "timeout": 0, "noerr": 0,
     "initialize": "", "finalize": "", "libs": [],
     "x": 400, "y": 400, "wires": [["sqlite_insert"]]},

    {"id": "sqlite_insert", "type": "sqlite", "z": TAB, "name": "写入历史",
     "mydb": SQLDB, "sqlquery": "prepared",
     "sql": "INSERT INTO readings (ts, node, rms, temp, state) VALUES (?, ?, ?, ?, ?)",
     "x": 630, "y": 400, "wires": [[]]},

    {"id": "inject_hist", "type": "inject", "z": TAB, "name": "每 30 秒查询",
     "props": [{"p": "payload"}], "repeat": "30", "crontab": "", "once": True,
     "onceDelay": "3", "topic": "", "payload": "", "payloadType": "date",
     "x": 200, "y": 520, "wires": [["func_hist"]]},

    {"id": "func_hist", "type": "function", "z": TAB, "name": "准备查询",
     "func": FUNC_HIST, "outputs": 1, "timeout": 0, "noerr": 0,
     "initialize": "", "finalize": "", "libs": [],
     "x": 400, "y": 520, "wires": [["sqlite_query"]]},

    {"id": "sqlite_query", "type": "sqlite", "z": TAB, "name": "查询最近 200 条",
     "mydb": SQLDB, "sqlquery": "prepared",
     "sql": "SELECT ts, rms FROM readings ORDER BY id DESC LIMIT 200",
     "x": 620, "y": 520, "wires": [["func_chart"]]},

    {"id": "func_chart", "type": "function", "z": TAB, "name": "转成曲线数据",
     "func": FUNC_CHART, "outputs": 1, "timeout": 0, "noerr": 0,
     "initialize": "", "finalize": "", "libs": [],
     "x": 820, "y": 520, "wires": [["chart_hist"]]},

    {"id": "chart_hist", "type": "ui_chart", "z": TAB, "name": "数据库历史曲线",
     "group": "grp_db", "order": 1, "width": "12", "height": "8",
     "label": "SQLite 历史（最近 200 条）", "chartType": "line", "legend": False,
     "xformat": "HH:mm:ss", "interpolate": "linear", "nodata": "数据库暂无数据",
     "ymin": "0", "ymax": "4",
     "removeOlder": "0", "removeOlderPoints": "", "removeOlderUnit": "3600",
     "cutout": 0, "useOneColor": True, "useUTC": False,
     "colors": ["#ff7f0e"], "outputs": 1, "x": 1050, "y": 520, "wires": [[]]},

    {"id": "file_1", "type": "file", "z": TAB, "name": "原始消息记录",
     "filename": LOG, "filenameType": "str", "appendNewline": True,
     "createDir": False, "overwriteFile": "false", "encoding": "none",
     "x": 400, "y": 300, "wires": []},

    {"id": "debug_1", "type": "debug", "z": TAB, "name": "调试输出",
     "active": True, "tosidebar": True, "console": False, "tostatus": False,
     "complete": "payload", "targetType": "msg",
     "statusVal": "", "statusType": "auto", "x": 400, "y": 350, "wires": []},
]

req = urllib.request.Request(
    "http://127.0.0.1:1880/flows",
    data=json.dumps(flows).encode("utf-8"),
    method="POST",
    headers={"Content-Type": "application/json",
             "Node-RED-Deployment-Type": "full"},
)
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        print("POST /flows ->", r.status, "| 节点数:", len(flows))
except urllib.error.HTTPError as e:
    print("HTTPError:", e.code, e.read().decode()[:400])
except Exception as e:
    print("失败:", e)
