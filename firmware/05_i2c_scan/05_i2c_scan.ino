// 步骤 3-1：I²C 总线扫描
// 目的：确认 MPU6050 接线正确、地址可读。不需要任何外部库。
// 接线：VCC→3V3  GND→GND  SDA→GPIO21(D21)  SCL→GPIO22(D22)  AD0→悬空
// 预期：发现设备 0x68（AD0 接高则为 0x69；若是 0x70 说明买到 MPU-6500）

#include <Wire.h>

void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);            // SDA=GPIO21, SCL=GPIO22
  Serial.println("\n=== I2C 总线扫描 ===");
  Serial.println("SDA=GPIO21  SCL=GPIO22  3.3V 供电");
}

void loop() {
  int found = 0;
  Serial.println("\n扫描中…");

  for (uint8_t addr = 1; addr < 127; addr++) {
    Wire.beginTransmission(addr);
    if (Wire.endTransmission() == 0) {
      Serial.printf("  发现设备: 0x%02X", addr);
      if (addr == 0x68)      Serial.print("   ← MPU6050 正常（AD0 悬空/接地）");
      else if (addr == 0x69) Serial.print("   ← MPU6050（AD0 接高）");
      else if (addr == 0x70) Serial.print("   ← MPU-6500，不是 MPU6050！");
      Serial.println();
      found++;
    }
  }

  if (found == 0) {
    Serial.println("  ✗ 没有发现任何 I2C 设备");
    Serial.println("    排查：① SDA/SCL 是否接反 ② VCC 是否 3.3V ③ 杜邦线是否插紧");
  }

  Serial.println("--- 3 秒后重新扫描 ---");
  delay(3000);
}
