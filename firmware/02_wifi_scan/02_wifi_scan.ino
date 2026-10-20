#include <WiFi.h>

void setup() {
  Serial.begin(115200);
  delay(500);
  WiFi.mode(WIFI_STA);
  WiFi.disconnect();
  delay(100);

  Serial.println("\n=== 扫描 2.4GHz 网络 ===");
  int n = WiFi.scanNetworks();
  if (n == 0) {
    Serial.println("没扫到任何网络");
  } else {
    Serial.printf("找到 %d 个：\n", n);
    for (int i = 0; i < n; i++) {
      Serial.printf("  %2d. %-26s %4d dBm  ch%-2d  %s\n",
        i + 1, WiFi.SSID(i).c_str(), WiFi.RSSI(i), WiFi.channel(i),
        WiFi.encryptionType(i) == WIFI_AUTH_OPEN ? "开放" : "加密");
    }
  }
  Serial.printf("\n本机 MAC: %s\n", WiFi.macAddress().c_str());
}

void loop() {}