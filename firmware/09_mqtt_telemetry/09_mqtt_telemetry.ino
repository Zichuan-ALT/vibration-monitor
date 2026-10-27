// 步骤 6：完整节点固件 v2 —— WiFi + MQTT + 振动测量 + 基线自学习 + 状态判定
//                     + NTP 对时 + 非阻塞重连 + 环形缓冲补发
// 库：MPU6050 (Electronic Cats) + arduinoFFT 2.x + PubSubClient
// 接线：VCC→3V3  GND→GND  SDA→GPIO21  SCL→GPIO22  AD0→悬空
//
// 工作流程
//   1. 连 WiFi、NTP 对时、连 MQTT
//   2. **基线自学习**：连续采 15 帧，算带内 RMS 的均值 μ 与标准差 σ
//      注意：自学习时设备必须处于"正常工况"（无配重），否则阈值会被学歪
//   3. 阈值 = μ+3σ（注意） / μ+6σ（报警）
//   4. 循环：测量 → 存入环形缓冲 → 若已连网则把缓冲全部发出
//   5. 迟滞判定：升级需连续 3 帧确认，降级需连续 5 帧确认
//
// v2 关键改进（store-and-forward）
//   · WiFi / MQTT 重连**不再阻塞**主循环，测量周期恒定为 2 s
//   · 网络断开时数据进入环形缓冲（60 条 = 约 2 分钟），恢复后按原顺序补发
//   · 每条数据携带**测量时刻的 NTP 时间戳**（ts，Unix 秒），
//     因此补发的数据在服务端仍能落在正确的时间位置
//   · 缓冲区满则丢弃最旧一条（记录丢弃计数）
//
// 上报主题：smartcampus/vib/{NODE_ID}/state
// 上报内容：
//   {"node","seq","ts","rms","peak","fpeak","tchip","state","warn","alarm"}
//
// 串口命令
//   r  重新学习基线
//   b  打印缓冲区与网络状态
//
// 现场使用顺序
//   上电 → 等自学习完成（约 20 秒）→ 正常工况运行 → 需要演示时加配重

#include <Wire.h>
#include <WiFi.h>
#include <PubSubClient.h>
#include <MPU6050.h>
#include <arduinoFFT.h>
#include <OneWire.h>
#include <DallasTemperature.h>
#include <SPI.h>
#include <SD.h>
#include <time.h>

// ================= 配置 =================
const char* SSID     = "XiXiaomi";
const char* PASS     = "12345sun.";
const char* BROKER   = "10.77.48.14";
const int   MQTT_PORT = 1883;
const char* NODE_ID  = "esp32-01";
const char* TOPIC_STATE = "smartcampus/vib/esp32-01/state";

// NTP
const char* NTP1 = "ntp.aliyun.com";
const char* NTP2 = "cn.pool.ntp.org";

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
#define BUF_SIZE     60      // 环形缓冲容量（60 条 × 2 s ≈ 2 分钟）
#define WIFI_RETRY_MS  10000UL
#define MQTT_RETRY_MS  15000UL   // 重连尝试本身是阻塞的（约 5 s TCP 超时），
                                 // 间隔放长可减少对采样节奏的影响
#define MIN_BASE_MU    0.010 // 基线下限 g：低于此值说明设备未运行或传感器脱落
                             // 若在此状态下学习，会把噪声底当正常运行水平，
                             // 设备一启动就会误报警，因此必须拒绝该基线

// ================= 数据结构 =================
struct Record {
  uint32_t seq;
  uint32_t ts;        // 测量时刻 Unix 秒（NTP），未对时为 0
  float rms, peak, fpeak, tchip, temp;
  uint8_t level;      // 0 正常 1 注意 2 报警
};

// ================= 全局 =================
MPU6050 mpu;

// 表面温度：DS18B20（单总线 GPIO5，模块板载 4.7k 上拉，无需外接）
#define ONE_WIRE_BUS 5
OneWire           oneWire(ONE_WIRE_BUS);
DallasTemperature ds18b20(&oneWire);
bool  tempOk = false;

// microSD（SPI：CLK=18 MISO=19 MOSI=23 CS=13，VCC 必须 5V）
#define SD_CS        13
#define LOG_PATH     "/vib_log.csv"
#define WAVE_DIR     "/waves"
bool  sdOk = false;
File  logFile;
bool  logHasHeader = false;

int16_t rawBuf[3][N];
double  vReal[N], vImag[N];
ArduinoFFT<double> FFT(vReal, vImag, N, (double)FS);

WiFiClient   espClient;
PubSubClient mqtt(espClient);

double  baseMu = 0, baseSigma = 0;
double  thrWarn = 0, thrAlarm = 0;
bool    learned = false;

bool     ntpSynced = false;
uint32_t lastWifiTry = 0, lastMqttTry = 0;

Record   ring[BUF_SIZE];
uint16_t ringHead = 0;    // 写入位置
uint16_t ringCount = 0;   // 当前条数
uint32_t dropped = 0;     // 缓冲满导致丢弃的条数
uint32_t resent = 0;      // 补发条数

uint32_t seq = 0;
double  lastRms = 0, lastPeak = 0, lastFreq = 0;
float   lastTemp = 0;     // MPU6050 芯片温度（参考）
float   lastSurf = NAN;   // DS18B20 表面温度（被测体）
uint8_t lastBestAxis = 0; // 本帧最强轴（0=ax 1=ay 2=az）
uint8_t stLevel = 0;
uint8_t stUp = 0, stDown = 0;

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

void measureOnce() {
  sampleBurst();
  double bestRms = -1;
  lastBestAxis = 0;
  for (uint8_t a = 0; a < 3; a++) {
    double r, ba, bf;
    analyseAxis(a, r, ba, bf);
    if (r > bestRms) { bestRms = r;  lastRms = r;  lastPeak = ba;  lastFreq = bf;  lastBestAxis = a; }
  }

  // 振动芯片温度（MPU6050 die 温度，仅作参考）
  lastTemp = mpu.getTemperature() / 340.0 + 36.53;

  // 被测体表面温度（DS18B20，12 位分辨率 = 0.0625 ℃）
  if (tempOk) {
    ds18b20.requestTemperatures();
    float t = ds18b20.getTempCByIndex(0);
    if (t > -50.0f && t < 125.0f) lastSurf = t;   // 越界则保留上次有效值
  }
}

// ================= 时间 =================
uint32_t nowEpoch() {
  time_t t = time(nullptr);
  return (t > 1600000000UL) ? (uint32_t)t : 0;   // 未对时返回 0
}

void syncNTP() {
  configTime(0, 0, NTP1, NTP2);        // 只取 UTC，时区由服务端处理
  Serial.print("NTP: 对时中 ");
  uint32_t t0 = millis();
  while (millis() - t0 < 8000) {
    if (nowEpoch() > 0) { ntpSynced = true; break; }
    Serial.print(".");  delay(500);
  }
  Serial.println();
  if (ntpSynced) {
    time_t t = time(nullptr);
    struct tm tmv;
    gmtime_r(&t, &tmv);
    Serial.printf("NTP OK  UTC %04d-%02d-%02d %02d:%02d:%02d\n",
                  tmv.tm_year + 1900, tmv.tm_mon + 1, tmv.tm_mday,
                  tmv.tm_hour, tmv.tm_min, tmv.tm_sec);
  } else {
    Serial.println("NTP 失败，时间戳将置 0（服务端回退为接收时刻）");
  }
}

// ================= 网络（全部非阻塞）=================
void ensureWiFi() {
  if (WiFi.status() == WL_CONNECTED) return;
  uint32_t now = millis();
  if (now - lastWifiTry < WIFI_RETRY_MS) return;
  lastWifiTry = now;
  WiFi.disconnect();
  WiFi.begin(SSID, PASS);
  Serial.printf("[WiFi] 重连中… (每 %lu 秒尝试一次)\n", WIFI_RETRY_MS / 1000);
}

void ensureMQTT() {
  if (WiFi.status() != WL_CONNECTED) return;
  if (mqtt.connected()) return;
  uint32_t now = millis();
  if (now - lastMqttTry < MQTT_RETRY_MS) return;
  lastMqttTry = now;
  Serial.printf("[MQTT] 连接 %s:%d …\n", BROKER, MQTT_PORT);
  if (mqtt.connect(NODE_ID)) {
    Serial.println("[MQTT] 已连接");
  } else {
    Serial.printf("[MQTT] 失败 rc=%d\n", mqtt.state());
  }
}

// ================= 环形缓冲 =================
void ringPush(const Record &r) {
  ring[ringHead] = r;
  ringHead = (ringHead + 1) % BUF_SIZE;
  if (ringCount < BUF_SIZE) ringCount++;
  else dropped++;                        // 满则覆盖最旧一条
}

bool ringPop(Record &out) {
  if (ringCount == 0) return false;
  uint16_t tail = (ringHead + BUF_SIZE - ringCount) % BUF_SIZE;
  out = ring[tail];
  ringCount--;
  return true;
}

const char* levelName(uint8_t lv) {
  if (!learned) return "uncalibrated";     // 未建立基线时不作判定
  return (lv == 2) ? "alarm" : (lv == 1) ? "warn" : "normal";
}

bool publishRecord(const Record &r) {
  char payload[340];
  char tempStr[16];
  if (isnan(r.temp)) strcpy(tempStr, "null");
  else                snprintf(tempStr, sizeof(tempStr), "%.2f", r.temp);

  snprintf(payload, sizeof(payload),
           "{\"node\":\"%s\",\"seq\":%lu,\"ts\":%lu,"
           "\"rms\":%.4f,\"peak\":%.4f,\"fpeak\":%.1f,"
           "\"temp\":%s,\"tchip\":%.1f,"
           "\"state\":\"%s\",\"warn\":%.4f,\"alarm\":%.4f}",
           NODE_ID, (unsigned long)r.seq, (unsigned long)r.ts,
           r.rms, r.peak, r.fpeak,
           tempStr, r.tchip,
           levelName(r.level), thrWarn, thrAlarm);
  return mqtt.publish(TOPIC_STATE, payload, false);
}

// 把缓冲里的数据按顺序全部发出去
void drainRing() {
  if (!mqtt.connected()) return;
  Record r;
  uint16_t sent = 0;
  while (ringCount > 0 && mqtt.connected()) {
    uint16_t tail = (ringHead + BUF_SIZE - ringCount) % BUF_SIZE;
    Record peek = ring[tail];
    if (!publishRecord(peek)) break;      // 发失败就留着，下轮再试
    ringPop(r);
    sent++;
    if (ringCount > 0 && sent >= 20) { mqtt.loop(); delay(10); }
  }
  if (sent > 0) {
    resent += sent;
    if (sent > 1) {   // 只有真正存在积压时才提示，正常上报不打印
      Serial.printf("[补发] %u 条已发出，缓冲剩余 %u 条\n", sent, ringCount);
    }
  }
}

// ================= 本地存储（microSD）=================
void sdInit() {
  SPI.begin(18, 19, 23, 13);
  sdOk = SD.begin(SD_CS, SPI, 4000000);
  if (!sdOk) {
    Serial.println("[SD] 未挂载（VCC 需 5V、卡需 FAT32），本地记录关闭");
    return;
  }
  Serial.printf("[SD] 挂载成功  %.1f GB\n", SD.cardSize() / (1024.0 * 1024.0 * 1024.0));

  if (!SD.exists(WAVE_DIR)) SD.mkdir(WAVE_DIR);

  logFile = SD.open(LOG_PATH, FILE_APPEND);
  if (!logFile) { Serial.println("[SD] 日志文件打开失败"); sdOk = false; return; }

  if (logFile.size() == 0) {
    logFile.println("ts_epoch,iso_utc,seq,rms,peak,fpeak,temp_surface,tchip_c,state");
  }
  Serial.printf("[SD] 日志 %s（%u 字节）\n", LOG_PATH, (unsigned)logFile.size());
}

void sdDumpWave(uint32_t ts) {
  if (!sdOk) return;
  char path[48];
  snprintf(path, sizeof(path), "%s/wave_%lu.csv", WAVE_DIR, (unsigned long)ts);
  File w = SD.open(path, FILE_WRITE);
  if (!w) { Serial.println("[SD] 波形文件创建失败"); return; }
  const char* nm[3] = {"ax", "ay", "az"};
  w.printf("# axis=%s rms=%.4f peak=%.4f fpeak=%.1f ts=%lu\n",
           nm[lastBestAxis], lastRms, lastPeak, lastFreq, (unsigned long)ts);
  w.println("index,raw_lsb");
  for (uint16_t i = 0; i < N; i++) w.printf("%u,%d\n", i, rawBuf[lastBestAxis][i]);
  w.close();
  Serial.printf("[SD] 已保存报警波形 %s（%d 点）\n", path, N);
}

void sdLog(uint32_t ts, uint32_t sq, bool alarmEdge) {
  if (!sdOk || !logFile) return;

  char iso[24] = "";
  if (ts > 1600000000UL) {
    time_t t = (time_t)ts;
    struct tm tmv;
    gmtime_r(&t, &tmv);
    strftime(iso, sizeof(iso), "%Y-%m-%dT%H:%M:%SZ", &tmv);
  }

  logFile.printf("%lu,%s,%lu,%.4f,%.4f,%.1f,%.2f,%.1f,%s\n",
                 (unsigned long)ts, iso, (unsigned long)sq,
                 lastRms, lastPeak, lastFreq,
                 isnan(lastSurf) ? -999.0 : lastSurf,
                 lastTemp, levelName(stLevel));
  logFile.flush();                 // 每条都刷盘，掉电最多丢一条

  if (alarmEdge) sdDumpWave(ts);   // 刚进入报警状态时保存原始波形
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

  // ---- 基线下限保护 ----
  // 设备停止时 RMS 只剩传感器噪声底（约 0.005 g）。若在此状态下学习，
  // 阈值会被压到 0.008 g 量级，设备一旦正常运转（约 0.05 g）立刻误报警。
  if (baseMu < MIN_BASE_MU) {
    learned = false;
    Serial.printf("\n!! 基线 mu = %.4f g 低于下限 %.3f g\n", baseMu, MIN_BASE_MU);
    Serial.println("!! 判定为：设备未运行 / 传感器脱落 / 未接触被测体");
    Serial.println("!! 已拒绝该基线，不启用报警判定");
    Serial.println("!! 请让设备正常运转后，串口发送 r 重新学习\n");
    return;
  }

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
void updateState() {
  uint8_t target = 0;
  if (lastRms >= thrAlarm)      target = 2;
  else if (lastRms >= thrWarn)  target = 1;

  if (target > stLevel)       { stUp++;   stDown = 0; }
  else if (target < stLevel)  { stDown++; stUp   = 0; }
  else                        { stUp = 0;  stDown = 0; }

  if (stUp >= 3)   { stLevel = target; stUp = 0; }
  if (stDown >= 5) { stLevel = target; stDown = 0; }
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
  bool mpuOk = mpu.testConnection();      // 真实检测，不写死

  // DS18B20 表面温度
  ds18b20.begin();
  ds18b20.setResolution(12);              // 0.0625 ℃
  tempOk = (ds18b20.getDeviceCount() > 0);

  Serial.println("\n=== 8.1 振动状态监测节点 v2（store-and-forward + 双传感）===");
  Serial.printf("节点 %s   上报 %s\n", NODE_ID, TOPIC_STATE);
  Serial.printf("MPU6050: %s（WHO_AM_I 原始 0x68）   DS18B20: %s（%d 个器件）\n",
                mpuOk ? "正常" : "未响应",
                tempOk ? "正常" : "未检测到", ds18b20.getDeviceCount());

  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.begin(SSID, PASS);

  Serial.print("WiFi: 连接 ");
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 20000) {
    Serial.print(".");  delay(500);
  }
  Serial.println();
  if (WiFi.status() == WL_CONNECTED)
    Serial.printf("WiFi OK   IP=%s  RSSI=%d\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI());
  else
    Serial.println("WiFi 首次连接未成功，将在后台继续重试（不阻塞测量）");

  if (WiFi.status() == WL_CONNECTED) syncNTP();

  mqtt.setServer(BROKER, MQTT_PORT);
  mqtt.setBufferSize(1024);
  ensureMQTT();

  sdInit();          // 本地存储（失败不影响联网上报）

  learnBaseline();
}

void loop() {
  // ---- 先测量，保证采样节奏（即使随后重连阻塞，本周期的数据也已经采到）----
  measureOnce();
  seq++;
  uint8_t prevLevel = stLevel;
  if (learned) updateState();
  bool alarmEdge = (learned && prevLevel != 2 && stLevel == 2);

  Record rec;
  rec.seq   = seq;
  rec.ts    = nowEpoch();
  rec.rms   = lastRms;
  rec.peak  = lastPeak;
  rec.fpeak = lastFreq;
  rec.tchip = lastTemp;
  rec.temp  = lastSurf;
  rec.level = stLevel;
  ringPush(rec);

  // ---- 本地记录（SD 卡）：每帧一行，报警瞬间另存原始波形 ----
  sdLog(rec.ts, seq, alarmEdge);

  // ---- 再处理网络（非阻塞检查；连接尝试本身可能阻塞数秒，但已不影响本周期的采样）----
  ensureWiFi();
  ensureMQTT();
  if (mqtt.connected()) {
    mqtt.loop();
    drainRing();
  }

  Serial.printf("RMS=%.4f 峰值=%.4f@%.1fHz  表面=%.2fC 芯片=%.1fC 状态=%s 缓冲=%u%s\n",
                lastRms, lastPeak, lastFreq, lastSurf, lastTemp,
                levelName(stLevel),
                ringCount, mqtt.connected() ? "" : "  [离线]");

  // ---- 串口命令 ----
  if (Serial.available()) {
    char c = Serial.read();
    if (c == 'r') { learnBaseline(); }
    if (c == 'b') {
      Serial.printf("[状态] WiFi=%d MQTT=%d NTP=%d 缓冲=%u 丢弃=%lu 补发=%lu\n",
                    (int)WiFi.status(), mqtt.connected(), ntpSynced,
                    ringCount, (unsigned long)dropped, (unsigned long)resent);
      Serial.printf("[SD]   %s  已用 %.1f MB\n",
                    sdOk ? "已挂载" : "未挂载",
                    SD.usedBytes() / (1024.0 * 1024.0));
    }
    if (c == 'l') {                       // 查看本地日志
      if (!sdOk) { Serial.println("[SD] 未挂载"); }
      else {
        File r = SD.open(LOG_PATH, FILE_READ);
        if (!r) { Serial.println("[SD] 日志文件不存在"); }
        else {
          uint32_t sz = r.size();
          Serial.printf("[SD] %s  %lu 字节\n", LOG_PATH, (unsigned long)sz);
          // 跳到末尾附近，打印最后若干行
          uint32_t start = (sz > 600) ? (sz - 600) : 0;
          r.seek(start);
          r.readStringUntil('\n');            // 丢弃可能被截断的半行
          uint8_t shown = 0;
          while (r.available() && shown < 10) {
            String line = r.readStringUntil('\n');
            Serial.print("  | ");
            Serial.println(line);
            shown++;
          }
          r.close();
        }
      }
    }
  }

  delay(PUB_PERIOD - 1024);
}
