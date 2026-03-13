#include <SPI.h>
#include <MFRC522.h>

#define RST_PIN         9          
#define SS_PIN          10         

MFRC522 mfrc522(SS_PIN, RST_PIN);  // Create MFRC522 instance

void setup() {
  Serial.begin(9600);   // Initialize serial communications
  SPI.begin();          // Init SPI bus
  mfrc522.PCD_Init();   // Init MFRC522 
  Serial.println("Scan an RFID tag...");
}

void loop() {
  // Look for new cards
  if ( ! mfrc522.PICC_IsNewCardPresent()) {
    return;
  }

  // Select one of the cards
  if ( ! mfrc522.PICC_ReadCardSerial()) {
    return;
  }

  // Combine the ID bytes into a single String
  String tagID = "";
  for (byte i = 0; i < mfrc522.uid.size; i++) {
    tagID += String(mfrc522.uid.uidByte[i] < 0x10 ? "0" : "");
    tagID += String(mfrc522.uid.uidByte[i], HEX);
  }

  tagID.toUpperCase();
  
  // This is what your Flask app will "read"
  Serial.println(tagID);

  // Halt PICC and stop encryption on PCD
  mfrc522.PICC_HaltA();
  mfrc522.PCD_StopCrypto1();
  
  delay(1000); // Wait a second before next read
}