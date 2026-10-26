// 步骤 6：完整节点固件 —— WiFi + MQTT + 振动测量 + 基线自学习 + 状态判定
// 库：MPU6050 (Electronic Cats) + arduinoFFT 2.x + PubSubClient
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// 工作流程
//   1. 连 WiFi、连 MQTT
//   2. **基线自学习**：连续采 15 帧，算带内 RMS 的均值 μ 与标准差 σ
//      注意：自学习时设备必须处于"正常工况"（无配重），否则阈值会被学歪
//   3. 阈值 = μ+3σ（注意） / μ+6σ（报警）
//   4. 循环测量并上报，采用迟滞判定：
//        升级需连续 3 帧确认，降级需连续 5 帧确认
//
// 上报主题：smartcampus/vib/{NODE_ID}/state
// 上报内容：
//   {"node","seq","rms","peak","fpeak","tchip","state","warn","alarm"}
//
// 现场使用顺序
//   上电 → 等自学习完成（约 20 秒）→ 正常工况运行 → 需要演示时加配重
//   重新学习：串口发 'r' 或重新上电

#include <Wire.h>
#include <WiFi.h>
#include <PubSubClient.h>
#include <MPU6050.h>
#include <arduinoFFT.h>

// ================= 配置 =================
const char* SSID     = "XiXiaomi";
const char* PASS     = "12345sun.";
const char* BROKER   = "10.77.48.14";
const int   MQTT_PORT = 1883;
const char* NODE_ID  = "esp32-01";
const char* TOPIC_STATE = "smartcampus/vib/esp32-01/state";

// ================= 采样参数 =================
#define N            1024
#define FS           1000.0
#define LSB_PER_G    8192.0
#define BAND_LO      10.0
#define BAND_HI      260.0
#define BASE_LO      20.0
#define BASE_HI      45.0

#define LEARN_FRAMES 15      // 自学习帧数（约 16 s）
#define PUB_PERIOD   2000    // 上报周期 ms（含测量时间）

// ================= 全局 =================
MPU6050 mpu;
int16_t rawBuf[3][N];
double  vReal[N], vImag[N];
ArduinoFFT<double> FFT(vReal, vImag, N, (double)FS);

WiFiClient   espClient;
PubSubClient mqtt(espClient);

double  baseMu = 0, baseSigma = 0;      // 基线统计
double  thrWarn = 0, thrAlarm = 0;      // 判定阈值
bool    learned = false;

uint32_t seq = 0;
double  lastRms = 0, lastPeak = 0, lastFreq = 0;
float   lastTemp = 0;
uint8_t stLevel = 0;                    // 0 正常 1 注意 2 报警
uint8_t stUp = 0, stDown = 0;           // 迟滞计数

// ================= 测量 =================
uint32_t sampleBurst() {
  uint32_t t0 = micros();
  uint32_t next = t0;
  const uint32_t periodUs = (uint32_t)(1000000.0 / FS + 0.5);
  for (uint16_t i = 0; i < N; i++) {
    while ((int32_t)(micros() - next) < 0) { }
    next += periodUs;
    int16_t ax, ay, az;
    mpu.getAcceleration(&ax, &ay, &az);
    rawBuf[0][i] = ax;  rawBuf[1][i] = ay;  rawBuf[2][i] = az;
  }
  return micros() - t0;
}

void analyseAxis(uint8_t a, double &rmsBand, double &baseAmp, double &baseFreq) {
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
  double sumBand = 0;
  baseAmp = 0;  baseFreq = 0;
  for (uint16_t k = 1; k < N / 2; k++) {
    double f   = k * df;
    double amp = vReal[k] / (N / 2.0) * 2.0;
    if (f >= BAND_LO && f <= BAND_HI) sumBand += amp * amp / 2.0;
    if (f >= BASE_LO && f <= BASE_HI && amp > baseAmp) { baseAmp = amp; baseFreq = f; }
  }
  rmsBand = sqrt(sumBand);
}

// 测一帧，取三轴中带内 RMS 最大者
void measureOnce() {
  sampleBurst();
  double bestRms = -1;
  for (uint8_t a = 0; a < 3; a++) {
    double r, ba, bf;
    analyseAxis(a, r, ba, bf);
    if (r > bestRms) {
      bestRms = r;  lastRms = r;  lastPeak = ba;  lastFreq = bf;
    }
  }
  lastTemp = mpu.getTemperature() / 340.0 + 36.53;
}

// ================= 网络 =================
void connectWiFi() {
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.begin(SSID, PASS);
  Serial.printf("WiFi: 连接 %s ", SSID);
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 30000) {
    Serial.print(".");  delay(500);
  }
  Serial.println();
  if (WiFi.status() == WL_CONNECTED)
    Serial.printf("WiFi OK   IP=%s  RSSI=%d\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI());
  else
    Serial.println("WiFi 失败，将继续离线运行");
}

void connectMQTT() {
  mqtt.setServer(BROKER, MQTT_PORT);
  mqtt.setBufferSize(1024);
  if (mqtt.connected()) return;
  Serial.printf("MQTT: 连接 %s:%d ", BROKER, MQTT_PORT);
  if (mqtt.connect(NODE_ID)) Serial.println("OK");
  else Serial.printf("失败 rc=%d\n", mqtt.state());
}

// ================= 基线自学习 =================
void learnBaseline() {
  Serial.printf("\n=== 基线自学习（%d 帧，约 %d 秒）===\n", LEARN_FRAMES, LEARN_FRAMES + 1);
  Serial.println("请确认设备处于正常工况（无配重），不要触碰设备");

  double sum = 0, sum2 = 0;
  for (uint8_t i = 0; i < LEARN_FRAMES; i++) {
    measureOnce();
    sum += lastRms;
    sum2 += lastRms * lastRms;
    Serial.printf("  [%2d/%d] RMS = %.4f g\n", i + 1, LEARN_FRAMES, lastRms);
  }

  baseMu = sum / LEARN_FRAMES;
  double var = sum2 / LEARN_FRAMES - baseMu * baseMu;
  if (var < 0) var = 0;
  baseSigma = sqrt(var);

  // 防止 sigma 过小导致阈值过紧：至少取 mu 的 15%
  double floorSigma = baseMu * 0.15;
  if (baseSigma < floorSigma) baseSigma = floorSigma;

  thrWarn  = baseMu + 3 * baseSigma;
  thrAlarm = baseMu + 6 * baseSigma;
  learned = true;

  Serial.printf("基线: mu = %.4f g,  sigma = %.4f g\n", baseMu, baseSigma);
  Serial.printf("阈值: 注意 = %.4f g,  报警 = %.4f g\n", thrWarn, thrAlarm);
  Serial.println("自学习完成，开始监测\n");
}

// ================= 判定（带迟滞）=================
const char* updateState() {
  uint8_t target = 0;
  if (lastRms >= thrAlarm)      target = 2;
  else if (lastRms >= thrWarn)  target = 1;

  if (target > stLevel)       { stUp++;   stDown = 0; }
  else if (target < stLevel)  { stDown++; stUp   = 0; }
  else                        { stUp = 0;  stDown = 0; }

  if (stUp >= 3)   { stLevel = target; stUp = 0; }
  if (stDown >= 5) { stLevel = target; stDown = 0; }

  return (stLevel == 2) ? "alarm" : (stLevel == 1) ? "warn" : "normal";
}

// ================= 上报 =================
void publishState(const char* state) {
  char payload[288];
  snprintf(payload, sizeof(payload),
           "{\"node\":\"%s\",\"seq\":%lu,\"rms\":%.4f,\"peak\":%.4f,"
           "\"fpeak\":%.1f,\"tchip\":%.1f,\"state\":\"%s\","
           "\"warn\":%.4f,\"alarm\":%.4f}",
           NODE_ID, (unsigned long)seq, lastRms, lastPeak,
           lastFreq, lastTemp, state, thrWarn, thrAlarm);

  if (mqtt.connected() && mqtt.publish(TOPIC_STATE, payload, false))
    Serial.printf("[发送] %s\n", payload);
  else
    Serial.printf("[失败] MQTT 未连接  %s\n", payload);
}

// ================= 主流程 =================
void setup() {
  Serial.begin(115200);
  delay(300);
  Wire.begin(21, 22);
  Wire.setClock(400000);

  mpu.initialize();
  mpu.setFullScaleAccelRange(MPU6050_ACCEL_FS_4);
  mpu.setDLPFMode(MPU6050_DLPF_BW_256);

  Serial.println("\n=== 8.1 振动状态监测节点 ===");
  Serial.printf("节点 %s   上报 %s\n", NODE_ID, TOPIC_STATE);

  connectWiFi();
  connectMQTT();
  learnBaseline();
}

void loop() {
  if (WiFi.status() != WL_CONNECTED) connectWiFi();
  if (!mqtt.connected())             connectMQTT();
  mqtt.loop();

  measureOnce();
  seq++;
  const char* state = updateState();
  publishState(state);

  Serial.printf("      RMS=%.4f  峰值=%.4f@%.1fHz  芯片温=%.1fC  状态=%s\n",
                lastRms, lastPeak, lastFreq, lastTemp, state);

  // 串口发 'r' 重新学习基线
  if (Serial.available()) {
    char c = Serial.read();
    if (c == 'r') { learned = false; learnBaseline(); }
  }

  delay(PUB_PERIOD - 1024);
}
