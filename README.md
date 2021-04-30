# Rice Data Analytics & Visualization Bootcamp <br> VBA Homework - The VBA of Wall Street

## Description
Use VBA scripting to analyze real stock market data.

## Input data
* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

## Instructions

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

### CHALLENGES

1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Solution

Main approach to the solution of the above problem was to write a universal VBA script, which would automatically determine number of sheets present within a given Workbook and loop though the varying number of rows of data within each Sheet individually.

VBA script was developed using the Test Data set and was then ran over the Stock Data set. Below are the snapshots of the three of the Workbook sheets after the completion of VBA script:

**2014**

<img src="Images\img_multi_year_stock_data_2014.JPG" alt="2014 Multi Year Stock analysis" style="zoom:75%;" />

**2015**

<img src="Images\img_multi_year_stock_data_2015.JPG" alt="2015 Multi Year Stock analysis" style="zoom:75%;" />

**2016**

<center> <img src="Images\img_multi_year_stock_data_2016.JPG" alt="2016 Multi Year Stock analysis" style="zoom:75%;" /></center>

