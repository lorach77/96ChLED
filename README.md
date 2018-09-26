# 96ChLED
R Script to control the 96ChLED optiogenetical plate from Life Cell Instrument. It overcomes most of the issuses which comes with the commercial software.

## R Script
The R_96ChLED.R script reads the excel table 96ChLED.xlsx which contains the program in the sheet “Program”. The other sheets are information or can be used to “save/store” old programs. How to set up the excel file can be read in the the sheet info in the same excel file. Which excel file shall be read in is hard coded in the R script in line number 20.

To run the script the best is to just source the Script in R studio. The output shows the duration of the run:

    > source('~/.../R_96ChLED.R')
    [1] "Run started at:  2018-06-28 11:34:47"
    [1] "Run will end at: 2018-06-28 11:35:00"
    [1] "Duration: 0.2 min."
    [1] "Duration: 0 hour."

To ease the adaptation of the script it is commented in detail.

## Installation

 - R Script: r_96chled.r
 -Excel File: 96chled.xlsx

It is important that the computer has an RS-232 output, else you cannot connect it to the Touch device. Further it needs following software: 

 - excel to edit the .xlsx file (recommended)
 - R
 - R Studio (recommended)
 - R-Packages: dplyr, stringr, serial and readxl

Then you just connect the Touchdevice to the RS-232 port and switch on the Touch device. Now you can start the Script.
Serial Codes
