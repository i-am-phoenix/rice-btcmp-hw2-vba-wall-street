# Rice Data Analytics & Visualization Bootcamp <br> VBA Homework - The VBA of Wall Street

## Description
Use VBA scripting to analyze real stock market data.

## Input data
* [Test Data](Resources/alphabetical_testing_0pt1.xlsm) - Used this data set for script development.

* [Stock Data](Resources/Multiple_year_stock_data_0pt1.xlsm) - Final analysis was run on this data using teh developed script.

## Instructions

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Implement conditional formatting that will highlight positive change in green and negative change in red.

### CHALLENGES

1. Script should be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

2. Make the appropriate adjustments to the VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Solution

Main approach to the solution of the above problem was to write a universal VBA script, which would automatically determine number of sheets present within a given Workbook and loop though the varying number of rows of data within each Sheet individually.

In the developed script, FOR loops were leveraged to cycle through sheets within an Active workbook, as well as through rows within the given sheet. The IF statement was then used to determine:
* a list of unique stock ticker names;
* record and maintain first and last recorded data for a given ticker (assuming that the data was provided in the ascending time format)
* calculate total_stock_volume per ticker as cumulative sum of recorded stock volumes in time

A new FOR loop was then created to determine highest positive and highest negative change in stock price and report corresponding stock ticker name as well as actual value of change, together with teh greatest total stock volume.

VBA script was developed using the Test Data set and was then run over the Stock Data set. Below are the snapshots of the three of the Workbook sheets after the completion of VBA script:

**2014**

<img src="Images\img_multi_year_stock_data_2014.JPG" alt="2014 Multi Year Stock analysis" style="zoom:75%;" />

**2015**

<img src="Images\img_multi_year_stock_data_2015.JPG" alt="2015 Multi Year Stock analysis" style="zoom:75%;" />

**2016**

<center> <img src="Images\img_multi_year_stock_data_2016.JPG" alt="2016 Multi Year Stock analysis" style="zoom:75%;" /></center>

