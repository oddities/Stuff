#include <SPI.h>
#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>

#define SCREEN_WIDTH 128 // OLED display width, in pixels
#define SCREEN_HEIGHT 64 // OLED display height, in pixels

// Declaration for an SSD1306 display connected to I2C (SDA, SCL pins)
#define OLED_RESET     4 // Reset pin # (or -1 if sharing Arduino reset pin)
Adafruit_SSD1306 display(SCREEN_WIDTH, SCREEN_HEIGHT, &Wire, OLED_RESET);
void setup() {
Serial.begin(9600);
Serial.println("sys started");

// SSD1306_SWITCHCAPVCC = generate display voltage from 3.3V internally
display.begin(SSD1306_SWITCHCAPVCC, 0x3C); // Address 0x3D for 128x64
// display.display();

// Show initial display buffer contents on the screen --
// the library initializes this with an Adafruit splash screen.
  Serial.println("in side setup");
  display.clearDisplay();
  display.display();
   display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(0,7);
 display.println("Starting  the device ..");
 display.display();
  delay(2000);
  Serial.println("in side void loop");
 display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(5,15);
 display.println("Searching for Wi-Fi Connection");
 display.display();
 delay(3000);
 display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(0,1);
 display.println("System    Ready for Testing");
 display.display();
 delay(3000);
 display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(5,15);
 display.println("Access    Point Mode");
 display.display();
 delay(3000);
   display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(2,3);
 display.println("1)Field    Test");
 display.println("2)Product  Test");
 display.display();
 delay(3000);// Pause for 2 seconds
   Serial.begin(9600);
   Serial.println("kypd setup");
  //configure pin 2 as an input and enable the internal pull-up resistor
  pinMode(2, INPUT_PULLUP);
  pinMode(13, OUTPUT);
  pinMode(3,INPUT_PULLUP);
}
 void loop(){
  If (digitalRead(2)==LOW)
  {
     Serial.println("in side 1st if main");
    display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(2,3);
 display.println("Nh3:\n NO3:\n soil:\n Temp:");
 display.display();
 delay(3000);
 display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(2,3);
 display.println("data sent to cloud");
 display.display();
 delay(3000);
  }
  else if (digitalRead(3)==LOW)
  {
    Serial.println("fV loop");
    display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(2,3);
 display.println("1)Fruits\n2)Vegetables");
 display.display();
 delay(3000);
 If (digitalRead(3)==LOW)
  {
    Serial.println("Apple banana loop");
    display.clearDisplay();
 display.setTextSize(2);
 display.setTextColor(WHITE);
 display.setCursor(2,3);
 display.println("1)Apple\n2)Banana");
 display.display();
 delay(3000);
}
  }
  
