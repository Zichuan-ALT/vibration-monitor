// 步骤 5：台架振动测量（带限 RMS + 三轴自动选轴）
// 库：MPU6050 (Electronic Cats) + arduinoFFT 2.x
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// 2.0 版修正（重要）
//   只有 10~260 Hz 的带内成分才算"振动"。
//   低频谱峰（<10 Hz）来自桌面晃动、安装松动、手持抖动，必须剔除，
//   否则会把"整体在晃"误当成"设备振动"。这与方案里"分析带宽 10~260 Hz"一致。
//
//   因此 RMS 改为从单边幅度谱按 Parseval 定理计算：
//       RMS_band = sqrt( Σ_{f∈带内} A_k² / 2 )
//   同时输出 <10 Hz 的低频 RMS，用于诊断安装是否牢固。
//
// 预期（风扇 1700 RPM）：基频 28.3 Hz 处出现明显谱峰。

#include <Wire.h>
#include <MPU6050.h>
#include <arduinoFFT.h>

#define N            1024
#define FS           1000.0
#define LSB_PER_G    8192.0
#define BAND_LO      10.0        // 分析带宽下限
#define BAND_HI      260.0       // 分析带宽上限
#define BASE_LO      20.0        // 基频带下限（实测转速 2052 RPM -> 34.2 Hz，原 15~35 无余量）
#define BASE_HI      45.0        // 基频带上限（覆盖 1200~2700 RPM）
#define VEL_LO       10.0
#define VEL_HI       250.0

MPU6050 mpu;
int16_t rawBuf[3][N];

double vReal[N];
double vImag[N];
ArduinoFFT<double> FFT(vReal, vImag, N, (double)FS);

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

// 对一个轴做 FFT，返回带内 RMS、低频 RMS、基频带峰值与其频率
void analyseAxis(uint8_t a, double &rmsBand, double &rmsLow, double &baseAmp, double &baseFreq) {
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

  // Hann 窗相干增益 0.5 → 幅值需乘 2
  const double df = FS / N;
  double sumBand = 0, sumLow = 0;
  baseAmp = 0; baseFreq = 0;

  for (uint16_t k = 1; k < N / 2; k++) {
    double f   = k * df;
    double amp = vReal[k] / (N / 2.0) * 2.0;
    double p   = amp * amp / 2.0;          // 该频点的均方值

    if (f < BAND_LO)               sumLow  += p;
    if (f >= BAND_LO && f <= BAND_HI) sumBand += p;

    if (f >= BASE_LO && f <= BASE_HI && amp > baseAmp) {
      baseAmp = amp; baseFreq = f;
    }
  }
  rmsBand = sqrt(sumBand);
  rmsLow  = sqrt(sumLow);
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);
  Wire.setClock(400000);

  mpu.initialize();
  mpu.setFullScaleAccelRange(MPU6050_ACCEL_FS_4);
  mpu.setDLPFMode(MPU6050_DLPF_BW_256);

  Serial.println("\n=== 台架振动测量 v2（带限 RMS 10~260 Hz）===");
  Serial.printf("N=%d  fs=%.0f Hz  Δf=%.3f Hz  基频带 %.0f~%.0f Hz\n",
                N, FS, FS / N, BASE_LO, BASE_HI);
  Serial.println("低频 RMS 用于诊断安装是否牢固；只有带内 RMS 才算设备振动。\n");
}

void loop() {
  uint32_t dur = sampleBurst();

  double rmsBand[3], rmsLow[3], baseAmp[3], baseFreq[3];
  for (uint8_t a = 0; a < 3; a++) {
    analyseAxis(a, rmsBand[a], rmsLow[a], baseAmp[a], baseFreq[a]);
  }

  // 按带内 RMS 选最强轴
  int best = 0;
  for (uint8_t a = 1; a < 3; a++) if (rmsBand[a] > rmsBand[best]) best = a;
  const char* nm[3] = {"ax", "ay", "az"};

  Serial.printf("fs=%.1f Hz  带内RMS: ax=%.4f ay=%.4f az=%.4f g\n",
                (N - 1) * 1e6 / dur, rmsBand[0], rmsBand[1], rmsBand[2]);
  Serial.printf("低频RMS:  ax=%.4f ay=%.4f az=%.4f g  (<10Hz，越小说明装得越稳)\n",
                rmsLow[0], rmsLow[1], rmsLow[2]);
  Serial.printf("选中 %s  带内RMS=%.4f g   基频带峰值=%.4f g @ %.1f Hz\n",
                nm[best], rmsBand[best], baseAmp[best], baseFreq[best]);
  Serial.println("----------------------------------------------------------");
  delay(700);
}
