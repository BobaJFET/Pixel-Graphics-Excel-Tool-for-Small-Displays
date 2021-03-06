# Pixel-Graphics-Excel-Tool-for-Small-Displays
A very simple tool to draw pixel graphics in excel for small displays such as OLED or TFT. Uses openpyxl library. This reads from the excel file and outputs code to be used in Adafruit GFX library. Code could also be easily modified to suit any type of pixel display. Refer to the Adafruit article for background: https://learn.adafruit.com/adafruit-gfx-graphics-library/overview 

Adafruit GFX Library for arduino allows you to easily draw basic shapes, lines, and individual pixels to small displays such as a monochrome OLED. But drawing graphics with code alone is time consuming and not intuitive. This tool allows you to draw your graphics in an excel file by filling cells. This script will simply output the drawPixel() and drawLine() functions, line by line to be pasted into your arduino or micro-controller code. 

Note: This code is currently set up for monochrome displays, where every pixel is the same color. I'm working on another script that will handle color and monochrome. 

1.) Create an excel file, draw your graphic using the fill tool. Use black to draw pixels. Only erase pixels with the 'no fill' color, or it will not output correctly. 

2.) Save the excel file, enter the path of the file into the quotes in the code. 

3.) Run and it will output all of the drawPixel() and drawLine() calls for your graphic. 

4.) Copy and paste the drawPixel() and drawLine() calls into your arduino code. 
