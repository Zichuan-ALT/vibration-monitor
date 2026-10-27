// 步骤 7：DS18B20 温度传感器测试与标定
// 库：OneWire 2.3.8 + DallasTemperature 4.0.6
// 接线：VCC→3V3  GND→GND  DAT→GPIO5（模块板载 4.7k 上拉，无需外接）
//
// 本程序用于
//   ① 确认传感器被正确识别（读出 ROM 地址与分辨率）
//   ② 观察读数稳定性（连续打印，判断噪声大小）
//   ③ 冰水浴标定（0 ℃）与室温对比
//
// 接线顺序提醒：探头三线 → 绿色端子：黄→DAT、红→VCC、黑→GND

#include <OneWire.h>
#include <DallasTemperature.h>

#define ONE_WIRE_BUS  5      // GPIO5（板上标 D5）
#define SAMPLE_MS     1000   // 打印间隔

OneWire           oneWire(ONE_WIRE_BUS);
DallasTemperature sensors(&oneWire);

DeviceAddress addr;
bool  found = false;
int   devCount = 0;

void printAddress(DeviceAddress a) {
  for (uint8_t i = 0; i < 8; i++) {
    Serial.printf("%02X", a[i]);
    if (i < 7) Serial.print(":");
  }
  Serial.println();
}

void setup() {
  Serial.begin(115200);
  delay(300);

  Serial.println("\n=== DS18B20 温度传感器测试 ===");
  Serial.printf("数据引脚: GPIO%d\n", ONE_WIRE_BUS);

  sensors.begin();
  sensors.setResolution(12);          // 12 位 = 0.0625 ℃ 分辨率
  devCount = sensors.getDeviceCount();

  Serial.printf("发现设备数: %d\n", devCount);

  if (devCount == 0) {
    Serial.println("\n!! 未发现器件，请检查:");
    Serial.println("   1. 端子线序: 黄→DAT、红→VCC、黑→GND");
    Serial.println("   2. 顶部三针: DAT→D5、VCC→3V3、GND→GND");
    Serial.println("   3. 杜邦线是否插紧");
    Serial.println("   4. 模块是否板载 4.7k 上拉（本模块已确认有 472 电阻）");
    sensors.begin();                  // 反复尝试
  } else {
    found = true;
    Serial.println("器件 ROM 地址:");
    for (uint8_t i = 0; i < devCount && i < 4; i++) {
      if (sensors.getAddress(addr, i)) {
        Serial.print("  #");
        Serial.print(i);
        Serial.print("  ");
        printAddress(addr);
      }
    }
    Serial.printf("分辨率: %d 位（%.4f ℃/LSB）\n",
                  sensors.getResolution(), sensors.getResolution() == 12 ? 0.0625 : 0.5);
    Serial.println("\n等待首次转换…");
  }
}

void loop() {
  if (!found) {
    sensors.begin();
    devCount = sensors.getDeviceCount();
    if (devCount > 0) {
      found = true;
      Serial.printf("\n✅ 检测到 %d 个器件，开始读数\n", devCount);
      if (sensors.getAddress(addr, 0)) { Serial.print("ROM: "); printAddress(addr); }
    } else {
      Serial.println("仍在等待器件…");
      delay(2000);
      return;
    }
  }

  sensors.requestTemperatures();
  float t = sensors.getTempCByIndex(0);

  if (t == DEVICE_DISCONNECTED_C) {
    Serial.println("!! 读取失败：器件断开（DEVICE_DISCONNECTED_C）");
  } else if (t <= -100.0f) {
    Serial.printf("!! 异常读数 %.1f ℃：通常是 DAT 未接好或缺少上拉电阻\n", t);
  } else if (fabsf(t - 85.0f) < 0.01f) {
    Serial.printf("!! 读数为 85.0 ℃：这是上电默认值，说明转换未完成或供电不足\n");
  } else {
    Serial.printf("温度 = %6.2f ℃", t);
    // 简单判据提示，便于现场判断接线是否正确
    if (t > 15.0f && t < 30.0f)        Serial.println("   （室温范围，正常）");
    else if (fabsf(t) < 1.0f)          Serial.println("   （约 0 ℃，冰水浴标定点）");
    else if (t > 30.0f && t < 45.0f)   Serial.println("   （温热水，用于对比标定）");
    else                               Serial.println();
  }

  delay(SAMPLE_MS);
}
