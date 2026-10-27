// 步骤 8：microSD 卡测试（含自动格式化）
// 接线：CLK→GPIO18  MISO→GPIO19  MOSI→GPIO23  CS→GPIO13  VCC→5V  GND→GND
//
// ⚠️ VCC 必须接 5V（模块板载 3.3V 稳压）
// ⚠️ 本程序在挂载失败时会尝试【格式化】卡为 FAT32
//    —— 卡上的所有数据会被清空！仅对空卡使用

#include <SD.h>
#include <SPI.h>

#define SD_CS    13

bool mounted = false;
uint32_t usedSpeed = 0;

const char* typeName(uint8_t t) {
  switch (t) {
    case CARD_MMC:  return "MMC";
    case CARD_SD:   return "SDSC";
    case CARD_SDHC: return "SDHC";
    case CARD_NONE: return "NONE";
    default:        return "UNKNOWN";
  }
}

bool tryMount(uint32_t freq, bool doFormat) {
  Serial.printf("  %lu kHz%s … ", (unsigned long)(freq / 1000),
                doFormat ? " +格式化" : "");
  if (SD.begin(SD_CS, SPI, freq, "/sd", 5, doFormat)) {
    Serial.printf("成功（%s）\n", typeName(SD.cardType()));
    mounted = true;
    usedSpeed = freq;
    return true;
  }
  Serial.println("失败");
  return false;
}

void setup() {
  Serial.begin(115200);
  delay(500);

  Serial.println("\n\n=== microSD 测试（含自动格式化）===");
  Serial.println("接线: CLK→D18  MISO→D19  MOSI→D23  CS→D13  VCC→5V\n");

  SPI.begin(18, 19, 23, 13);

  // 第一步：正常挂载（不动卡上的数据）
  Serial.println("[1] 正常挂载:");
  tryMount(4000000, false) || tryMount(1000000, false) || tryMount(20000000, false);

  // 第二步：仍失败则尝试格式化（会清空卡）
  if (!mounted) {
    Serial.println("\n[2] 正常挂载失败，尝试格式化（卡上数据将被清空）…");
    Serial.println("    如果是新卡，这是安全的");
    delay(2000);
    tryMount(4000000, true) || tryMount(1000000, true);
  }

  if (!mounted) {
    Serial.println("\n!! 仍然失败。硬件问题的可能性增大:");
    Serial.println("   1. 四根信号线接错（重点查 MISO/MOSI 是否反了）");
    Serial.println("   2. 卡没插到底");
    Serial.println("   3. 模块或卡本身损坏");
    Serial.println("\n建议：换一张卡试试，或检查模块丝印与接线是否一一对应。");
    return;
  }

  Serial.printf("\n✅ 挂载成功（SPI %lu kHz）\n", (unsigned long)(usedSpeed / 1000));
  Serial.printf("卡类型:  %s\n", typeName(SD.cardType()));
  Serial.printf("容量:    %.1f GB\n", SD.cardSize() / (1024.0 * 1024.0 * 1024.0));
  Serial.printf("总空间:  %.1f GB\n", SD.totalBytes() / (1024.0 * 1024.0 * 1024.0));
  Serial.printf("已用:    %.1f MB\n", SD.usedBytes() / (1024.0 * 1024.0));

  // 写读测试
  Serial.println("\n--- 写入测试 ---");
  File f = SD.open("/selftest.txt", FILE_WRITE);
  if (!f) {
    Serial.println("!! 无法创建文件（卡可能写保护）");
    return;
  }
  f.println("DSH SD self-test");
  f.printf("speed_khz=%lu\n", (unsigned long)(usedSpeed / 1000));
  f.printf("size_gb=%.1f\n", SD.cardSize() / (1024.0 * 1024.0 * 1024.0));
  f.close();

  Serial.println("--- 读回测试 ---");
  f = SD.open("/selftest.txt", FILE_READ);
  if (f) {
    while (f.available()) {
      Serial.print("  | ");
      Serial.println(f.readStringUntil('\n'));
    }
    f.close();
    SD.remove("/selftest.txt");
    Serial.println("\n✅ 读写均正常，测试文件已删除\n");
  } else {
    Serial.println("!! 读取失败\n");
  }
}

void loop() {
  static uint32_t last = 0;
  if (millis() - last > 5000) {
    last = millis();
    if (mounted) {
      Serial.printf("[心跳] %lu s   可用空间 %.1f GB\n",
                    (unsigned long)(millis() / 1000),
                    (SD.totalBytes() - SD.usedBytes()) / (1024.0 * 1024.0 * 1024.0));
    } else {
      Serial.println("[心跳] 未挂载");
    }
  }
}
