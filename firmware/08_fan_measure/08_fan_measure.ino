// 步骤 5：台架振动测量（带限 RMS + 三轴自动选轴 + 频谱导出）
// 库：MPU6050 (Electronic Cats) + arduinoFFT 2.x
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// v3 新增：串口收到任意字符时，把当前帧的频谱以 CSV 输出，便于上位机保存与绘图。
//
// 输出约定
//   正常运行时打印一行摘要
//   收到字符后追加打印：
//     #BEGIN,<轴向编号>,<RMS>,<峰值频率>,<峰值幅值>
//     <频率Hz>,<幅值g>   (逐行，0~500 Hz，步长 fs/N)
//     #END
//
// 说明
//   只有 10~260 Hz 的带内成分才算"振动"。低频谱峰（<10 Hz）来自桌面晃动、
//   安装松动、手持抖动，必须剔除。这与方案里"分析带宽 10~260 Hz"一致。
//   RMS 由单边幅度谱按 Parseval 定理计算：RMS = sqrt( Σ A_k² / 2 )

#include <Wire.h>
#include <MPU6050.h>
#include <arduinoFFT.h>

#define N            1024
#define FS           1000.0
#define LSB_PER_G    8192.0
#define BAND_LO      10.0
#define BAND_HI      260.0
#define BASE_LO      20.0
#define BASE_HI      45.0

MPU6050 mpu;
int16_t rawBuf[3][N];

double vReal[N];
double vImag[N];
ArduinoFFT<double> FFT(vReal, vImag, N, (double)FS);

double  lastAmp[N / 2];
uint8_t lastAxis = 0;
double  lastRms = 0, lastPeakAmp = 0, lastPeakFreq = 0;

uint32_t sampleBurst() {
  uint32_t t0 = micros();
  uint32_t next = t0;
  const uint32_t periodUs = (uint32_t)(1000000.0 / FS + 0.5);

  for (uint16_t i = 0; i < N; i++) {
    while ((int32_t)(micros() - next) < 0) { }
    next += periodUs;
    int16_t ax, ay, az;
    mpu.getAcceleration(&ax, &ay, &az);
    rawBuf[0][i] = ax;
    rawBuf[1][i] = ay;
    rawBuf[2][i] = az;
  }
  return micros() - t0;
}

void analyseAxis(uint8_t a, double &rmsBand, double &rmsLow,
                 double &baseAmp, double &baseFreq, double *ampOut) {
  double s = 0;
  for (uint16_t i = 0; i < N; i++) s += rawBuf[a][i];
  double mean = s / N;

  for (uint16_t i = 0; i < N; i++) {
    vReal[i] = (rawBuf[a][i] - mean) / LSB_PER_G;
    vImag[i] = 0.0;
  }
  FFT.windowing(FFTWindow::Hann, FFTDirection::Forward);
  FFT.compute(FFTDirection::Forward);
  FFT.complexToMagnitude();

  const double df = FS / N;
  double sumBand = 0, sumLow = 0;
  baseAmp = 0; baseFreq = 0;

  for (uint16_t k = 1; k < N / 2; k++) {
    double f   = k * df;
    double amp = vReal[k] / (N / 2.0) * 2.0;
    double p   = amp * amp / 2.0;

    if (ampOut) ampOut[k] = amp;
    if (f < BAND_LO)                  sumLow  += p;
    if (f >= BAND_LO && f <= BAND_HI) sumBand += p;
    if (f >= BASE_LO && f <= BASE_HI && amp > baseAmp) { baseAmp = amp; baseFreq = f; }
  }
  rmsBand = sqrt(sumBand);
  rmsLow  = sqrt(sumLow);
}

void dumpSpectrum() {
  Serial.printf("#BEGIN,%u,%.6f,%.2f,%.6f\n",
                lastAxis, lastRms, lastPeakFreq, lastPeakAmp);
  const double df = FS / N;
  for (uint16_t k = 1; k < N / 2; k++) {
    Serial.printf("%.3f,%.7f\n", k * df, lastAmp[k]);
  }
  Serial.println("#END");
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);
  Wire.setClock(400000);

  mpu.initialize();
  mpu.setFullScaleAccelRange(MPU6050_ACCEL_FS_4);
  mpu.setDLPFMode(MPU6050_DLPF_BW_256);

  Serial.println("\n=== 台架振动测量 v3（带限 RMS + 频谱导出）===");
  Serial.printf("N=%d  fs=%.0f Hz  Δf=%.3f Hz  基频带 %.0f~%.0f Hz\n",
                N, FS, FS / N, BASE_LO, BASE_HI);
  Serial.println("发送任意字符可导出当前帧频谱 CSV\n");
}

void loop() {
  uint32_t dur = sampleBurst();

  double rmsBand[3], rmsLow[3], baseAmp[3], baseFreq[3];
  for (uint8_t a = 0; a < 3; a++) {
    analyseAxis(a, rmsBand[a], rmsLow[a], baseAmp[a], baseFreq[a], nullptr);
  }

  int best = 0;
  for (uint8_t a = 1; a < 3; a++) if (rmsBand[a] > rmsBand[best]) best = a;

  // 只对最强轴保存频谱
  {
    double r, l, ba, bf;
    analyseAxis(best, r, l, ba, bf, lastAmp);
  }
  lastAxis     = best;
  lastRms      = rmsBand[best];
  lastPeakAmp  = baseAmp[best];
  lastPeakFreq = baseFreq[best];

  const char* nm[3] = {"ax", "ay", "az"};
  Serial.printf("fs=%.1f Hz  带内RMS: ax=%.4f ay=%.4f az=%.4f g\n",
                (N - 1) * 1e6 / dur, rmsBand[0], rmsBand[1], rmsBand[2]);
  Serial.printf("低频RMS:  ax=%.4f ay=%.4f az=%.4f g  (<10Hz)\n",
                rmsLow[0], rmsLow[1], rmsLow[2]);
  Serial.printf("选中 %s  带内RMS=%.4f g   基频带峰值=%.4f g @ %.1f Hz\n",
                nm[best], rmsBand[best], baseAmp[best], baseFreq[best]);

  if (Serial.available()) {
    while (Serial.available()) Serial.read();
    dumpSpectrum();
  }

  Serial.println("----------------------------------------------------------");
  delay(700);
}
