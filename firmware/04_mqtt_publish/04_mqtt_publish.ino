#include <WiFi.h>
#include <PubSubClient.h>

// ===== 网络配置 =====
const char* SSID    = "XiXiaomi";
const char* PASS    = "12345sun.";
const char* BROKER  = "10.77.48.14";
const int   PORT    = 1883;
const char* NODE_ID = "esp32-test";

WiFiClient espClient;
PubSubClient mqtt(espClient);

void connectWiFi() {
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.begin(SSID, PASS);
  Serial.printf("WiFi: 连接 %s ", SSID);
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 30000) {
    Serial.print(".");
    delay(500);
  }
  Serial.println();
  if (WiFi.status() == WL_CONNECTED) {
    Serial.printf("WiFi OK   IP=%s   RSSI=%d\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI());
  } else {
    Serial.printf("WiFi 失败  status=%d\n", WiFi.status());
  }
}

void connectMQTT() {
  mqtt.setServer(BROKER, PORT);
  mqtt.setBufferSize(1024);
  Serial.printf("MQTT: 连接 %s:%d ", BROKER, PORT);
  int tries = 0;
  while (!mqtt.connected() && tries < 10) {
    tries++;
    if (mqtt.connect(NODE_ID)) { Serial.println("OK"); return; }
    Serial.printf("失败 rc=%d ", mqtt.state());
    delay(1500);
  }
  Serial.println();
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Serial.println("\n=== MQTT 上报测试（模拟工况变化）===");
  connectWiFi();
  connectMQTT();
}

void loop() {
  if (WiFi.status() == WL_CONNECTED && !mqtt.connected()) connectMQTT();
  mqtt.loop();

  static uint32_t last = 0;
  static uint32_t seq  = 0;
  if (millis() - last >= 3000) {
    last = millis();
    seq++;

    // --- 模拟渐进式不平衡：90 秒一个周期，RMS 从 0.010 平滑升到 0.070 再回落 ---
    float phase = (millis() % 90000UL) / 90000.0f;      // 0 .. 1
    float rms   = 0.010f + 0.060f * (0.5f * (1.0f - cosf(2.0f * PI * phase)));
    float temp  = 26.0f + 8.0f * phase;                  // 温度随负载缓慢上升

    const char* state = (rms >= 0.060f) ? "alarm"
                      : (rms >= 0.030f) ? "warn" : "normal";

    char topic[96];
    snprintf(topic, sizeof(topic), "smartcampus/vib/%s/state", NODE_ID);

    char payload[256];
    snprintf(payload, sizeof(payload),
             "{\"node\":\"%s\",\"seq\":%lu,\"rms\":%.4f,\"temp\":%.1f,\"state\":\"%s\"}",
             NODE_ID, (unsigned long)seq, rms, temp, state);

    bool ok = mqtt.publish(topic, payload, false);
    Serial.printf("[%s] seq=%lu  %s\n", ok ? "发送" : "失败",
                  (unsigned long)seq, payload);
  }
}
