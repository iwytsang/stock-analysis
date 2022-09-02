# Refactoring Stock Analysis

## Overview of Project

### Purpose
The purpose of the analysis is to assist Steve with analyzing a entire dataset of stocks. Although the original code previously written for Steve excuted well for a dozen stocks, it may not run the most efficiently with thousands of stocks. With refactoring the code, Steve can analyze not just a dozen stocks, but thousands of stocks more quickly and efficiently.

## Results

The returns of all stocks worsened between 2017 and 2018 for all stocks except "TERP", which had a return of -7.2% in 2017 and -5.0% in 2018. Total daily volume of stocks traded in 2017 summed to 3,166,639,100 and in 2018 summed to 3,306,038,200.

In 2017, all stocks except "TERP" had positive returns from the starting price to ending price of the stock. 

![Stock_Results_2017](https://user-images.githubusercontent.com/108503112/188239681-bd2ddfcb-7901-4ada-808f-6be0d4a6b7dc.png)

In 2018, all stocks except "ENPH" and "RUN" had negative returns from the starting price to ending price of the stock.

![Stock_Results_2018](https://user-images.githubusercontent.com/108503112/188239704-908d4cb4-aa71-4a4a-8d90-77efa8d9cf83.png)

The refactored code ran much quicker than the original code written. As shown in the images below, the refactored code ran in 0.0546875 seconds for 2017 and 2018, compared with the original code which ran in 0.2734375 seconds for 2017 and 2018.

###Refactored Code Run-Times

![VBA_Challenge_2017](https://user-images.githubusercontent.com/108503112/188242281-a2943bec-a9d0-4228-be90-36dcfeb2667b.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/108503112/188242285-87790d9e-f490-4b45-a70c-3d53cabdcff2.png)

###Original Code Run-Times

![Original_Code_2017](https://user-images.githubusercontent.com/108503112/188242677-da29dd1f-ab16-4d43-8ee2-639053202364.png)
![Original_Code_2018](https://user-images.githubusercontent.com/108503112/188242681-0c27dccd-f8d5-4385-be52-5e4b0fc57291.png)

## Summary

###Advantages and Disadvantages of Refactoring VBA Code in General
Advantages of refactoring VBA code is that it can run more quickly and efficiently to analyze larger amounts of data in a dataset.

Disadvantages of refactoring VBA code is that while going through the refactoring process, the code can break and have many errors that are difficult and time-consuming to fix. 

###Advantages and Disadvantages of Refactored Stock Analysis

