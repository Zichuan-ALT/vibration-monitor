// 步骤 1：最小系统验证 —— 板载 LED 闪烁
void setup() {
  Serial.begin(115200);
  pinMode(2, OUTPUT);          // GPIO2 = 板载 LED
  Serial.println("ESP32 alive");
}
void loop() {
  digitalWrite(2, HIGH);
  delay(500);
  digitalWrite(2, LOW);
  delay(500);
}
