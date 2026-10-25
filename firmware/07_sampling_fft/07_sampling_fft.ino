// 步骤 4：1 kHz 采样 + FFT 特征提取
// 库：MPU6050 (Electronic Cats) + arduinoFFT 2.x
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// 本程序验证三件事：
//   ① 能否稳定做到 1 kSPS（采样 1024 点应为 1023 ms 左右）
//   ② 静态噪声底是多少（决定能测到多小的振动）
//   ③ FFT 与特征值（RMS / 峰值频率 / 基频带 / 速度有效值）是否正常
//
// 输出每次刷新，可与《实施方案》表 23 的参数对照。

#include <Wire.h>
#include <MPU6050.h>
#include <arduinoFFT.h>

#define N            1024
#define FS           1000.0
#define LSB_PER_G    8192.0      // ±4 g 量程
#define BAND_LO      15.0        // 基频带下限 Hz
#define BAND_HI      35.0        // 基频带上限 Hz
#define VEL_LO       10.0        // 速度积分下限 Hz
#define VEL_HI       250.0       // 速度积分上限 Hz

MPU6050  mpu;
double   vReal[N];
double   vImag[N];
int16_t  raw[N];

// arduinoFFT 2.x 是模板类：ArduinoFFT<T>，不是 arduinoFFT
ArduinoFFT<double> FFT(vReal, vImag, N, (double)FS);

// ---------- 采样：严格 1 kSPS ----------
uint32_t sampleBurst() {
  uint32_t t0   = micros();
  uint32_t next = t0;
  const uint32_t periodUs = (uint32_t)(1000000.0 / FS + 0.5);

  for (uint16_t i = 0; i < N; i++) {
    while ((int32_t)(micros() - next) < 0) { }   // 等到采样时刻
    next += periodUs;
    int16_t ax, ay, az;
    mpu.getAcceleration(&ax, &ay, &az);
    raw[i] = az;                                 // 本测试用 Z 轴
  }
  return micros() - t0;
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);
  Wire.setClock(400000);                         // 400 kHz，缩短读取耗时

  mpu.initialize();
  mpu.setFullScaleAccelRange(MPU6050_ACCEL_FS_4);
  mpu.setDLPFMode(MPU6050_DLPF_BW_256);          // 256 Hz 带宽

  Serial.println("\n=== 采样 + FFT 特征提取 ===");
  Serial.printf("N=%d   fs=%.0f Hz   Δf=%.3f Hz   分析带宽 0~%.0f Hz\n",
                N, FS, FS / N, FS / 2);
  Serial.printf("基频带 %.0f~%.0f Hz   速度积分 %.0f~%.0f Hz\n",
                BAND_LO, BAND_HI, VEL_LO, VEL_HI);
  Serial.println("首次采样中…\n");
}

void loop() {
  uint32_t dur = sampleBurst();

  // ---- 时域统计 ----
  double sum = 0;
  for (uint16_t i = 0; i < N; i++) sum += raw[i];
  double mean = sum / N;

  double sq = 0;
  for (uint16_t i = 0; i < N; i++) {
    double d = raw[i] - mean;
    sq += d * d;
  }
  double rms = sqrt(sq / N) / LSB_PER_G;          // AC 有效值 [g]
  double std = rms;                                // 静态时 RMS≈σ

  // ---- 去直流 + 加窗 ----
  for (uint16_t i = 0; i < N; i++) {
    vReal[i] = (raw[i] - mean) / LSB_PER_G;
    vImag[i] = 0.0;
  }
  FFT.windowing(FFTWindow::Hann, FFTDirection::Forward);
  FFT.compute(FFTDirection::Forward);
  FFT.complexToMagnitude();

  // ---- 峰值与基频带 ----
  double df    = FS / N;
  double fPeak = 0, aPeak = 0;                    // 全带峰值
  double fBand = 0, aBand = 0;                    // 基频带峰值
  double vSq   = 0;                               // 速度积分

  for (uint16_t k = 1; k < N / 2; k++) {
    double f = k * df;
    double a = vReal[k] / (N / 2.0) * 2.0;        // 幅值（Hann 窗近似恢复）

    if (f <= VEL_HI && a > aPeak) { aPeak = a; fPeak = f; }
    if (f >= BAND_LO && f <= BAND_HI && a > aBand) { aBand = a; fBand = f; }

    if (f >= VEL_LO && f <= VEL_HI) {
      double v = a / (2.0 * PI * f);              // 加速度谱 → 速度谱
      vSq += v * v;
    }
  }
  double vRms = sqrt(vSq / 2.0) * 1000.0;         // m/s → mm/s

  // ---- 采样率核算 ----
  double actualFs = (N - 1) * 1e6 / dur;

  // ---- 输出 ----
  Serial.printf("采样耗时 %lu ms   实际 fs = %.1f Hz  (误差 %+.2f%%，判据 ≤0.5%%)\n",
                (unsigned long)(dur / 1000), actualFs,
                (actualFs - FS) / FS * 100.0);
  Serial.printf("DC(均值) = %+.4f g\n", mean / LSB_PER_G);
  Serial.printf("噪声底 σ = %.5f g      3σ 门限 = %.5f g\n", std, 3 * std);
  Serial.printf("RMS(AC)  = %.5f g\n", rms);
  Serial.printf("全带峰值 = %.5f g @ %.1f Hz\n", aPeak, fPeak);
  Serial.printf("基频带峰值 = %.5f g @ %.1f Hz\n", aBand, fBand);
  Serial.printf("速度有效值 = %.2f mm/s  (%.0f~%.0f Hz)\n",
                vRms, VEL_LO, VEL_HI);
  Serial.println("---------------------------------------------");
  delay(800);
}
