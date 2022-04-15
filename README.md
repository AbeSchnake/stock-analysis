# Stock Analysis Challenge Write-up
## Overview
The goal of this project was to analyze the performace of various stocks in the years 2017 and 2018, and to refactor the VBA code performing the analysis to make the code run faster.
## Results
Below is a comparison of the two output tables for 2017 and 2018, respectively. Every stock within the analysis performed significantly worse in 2018 than they did in 2017, with the one exception of RUN, which performed significantly better. Analyzing market trends that might explain the worse 2018 performance is beyond the scope of this project, and based on the information available my reccomendation would be to invest in RUN.
![2017 Performance](https://github.com/AbeSchnake/stock-analysis/blob/main/Resources/2017%20Stock%20Performance.png)
![2018 Performance](https://github.com/AbeSchnake/stock-analysis/blob/main/Resources/2018%20Stock%20Performance.png)

Below is a comparison of the runtimes for 2017 and 2018. These times are about one quarter of the runtime of the original code. This difference was achieved by using a Ticker Index variable to loop through the rows in the data tables only once, rather than looping through them for each different ticker. The outputs for each ticker were calculated using the following code pasted from my VBA editor with the comments removed:
        > tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
              
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                        
       End If
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
             tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
![2017 Runtime](https://github.com/AbeSchnake/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018 Runtime](https://github.com/AbeSchnake/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
