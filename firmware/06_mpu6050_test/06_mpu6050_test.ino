// 步骤 3-2：MPU6050 自检（对应 AC 里的传感器验收判据）
// 需要库：MPU6050 by Electronic Cats
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// 【重要】Electronic Cats 的库对 WHO_AM_I 做了归一化：
//   readBits(reg 0x75, bit 6, len 6) → 实际拿到的是 (WHO_AM_I >> 1) 的低 6 位
//   原厂 MPU6050 的 WHO_AM_I = 0x68  →  库返回 0x34
//   所以判据要用 0x34，不是 0x68。本程序同时读原始寄存器，便于留证。
//
// 三项判据：
//   ① 原始 WHO_AM_I = 0x68（库归一化值 0x34）
//   ② 静止时 |a| ≈ 1.000 g（零偏 ≤ ±0.2 g）
//   ③ 轻敲桌面，数值明显跳变

#include <Wire.h>
#include <MPU6050.h>

MPU6050 mpu;

const float LSB_PER_G = 8192.0;      // ±4 g 量程
const uint8_t ADDR = 0x68;           // AD0 悬空时的地址

// 直接读 WHO_AM_I(0x75) 原始值，绕过库的归一化
uint8_t readWhoAmIRaw() {
  Wire.beginTransmission(ADDR);
  Wire.write(0x75);
  if (Wire.endTransmission(false) != 0) return 0xFF;
  if (Wire.requestFrom(ADDR, (uint8_t)1) != 1) return 0xFE;
  return Wire.read();
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);

  Serial.println("\n=== MPU6050 自检 ===");
  mpu.initialize();
  mpu.setFullScaleAccelRange(MPU6050_ACCEL_FS_4);   // ±4 g

  uint8_t raw  = readWhoAmIRaw();
  uint8_t norm = mpu.getDeviceID();

  Serial.printf("WHO_AM_I 原始寄存器 = 0x%02X  （原厂 MPU6050 应为 0x68）\n", raw);
  Serial.printf("库归一化 ID          = 0x%02X  （库规定应为 0x34）\n", norm);
  Serial.printf("连接测试             = %s\n", mpu.testConnection() ? "OK" : "失败");

  if (raw == 0x68 && norm == 0x34)      Serial.println("判据① 通过：原厂 MPU6050");
  else if (raw == 0x70)                 Serial.println("判据① 不通过：这是 MPU-6500");
  else                                  Serial.println("判据① 存疑：ID 非常规，但若读数正常仍可用");

  Serial.println("\n判据②：静止时 |a| 应约 1.000 g（零偏 ≤ ±0.2 g）");
  Serial.println("判据③：轻敲桌面时数值应明显跳变\n");
}

void loop() {
  int16_t ax, ay, az, gx, gy, gz;
  mpu.getMotion6(&ax, &ay, &az, &gx, &gy, &gz);

  float fx = ax / LSB_PER_G;
  float fy = ay / LSB_PER_G;
  float fz = az / LSB_PER_G;
  float mag = sqrtf(fx * fx + fy * fy + fz * fz);

  Serial.printf("ax=%+7.3f ay=%+7.3f az=%+7.3f  |a|=%.3f g\n", fx, fy, fz, mag);
  delay(300);
}
