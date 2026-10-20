#include <WiFi.h>

const char* SSID = "XiXiaomi";        // ← 你的热点名
const char* PASS = "12345sun.";         // ← 你的密码

void setup() {
  Serial.begin(115200);
  delay(300);
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);                // 关省电，响应更快
  WiFi.begin(SSID, PASS);

  Serial.printf("\n连接 %s ", SSID);
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 30000) {
    Serial.print(".");
    delay(500);
  }
  Serial.println();

  if (WiFi.status() == WL_CONNECTED) {
    Serial.println("OK 连接成功");
    Serial.printf("ESP32 IP = %s\n", WiFi.localIP().toString().c_str());
    Serial.printf("网关     = %s\n", WiFi.gatewayIP().toString().c_str());
    Serial.printf("RSSI     = %d dBm\n", WiFi.RSSI());
    Serial.printf("MAC      = %s\n", WiFi.macAddress().c_str());
  } else {
    Serial.printf("失败，status = %d\n", WiFi.status());
  }
}

void loop() {
  static uint32_t t = 0;
  if (millis() - t > 5000) {
    t = millis();
    Serial.printf("[状态] %s  IP=%s  RSSI=%d\n",
      WiFi.status() == WL_CONNECTED ? "在线" : "离线",
      WiFi.localIP().toString().c_str(), WiFi.RSSI());
  }
}