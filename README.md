# Stock Analysis with VBA

## Overview of Project

The purpose of this project was to help my friend, Steve, automate some routine analysis in his first job as a financial advisor. The first goal was to write code that would take a group of stocks in a given year and determine 1) their total trading volume that year and 2) your yearly return had you held the stock the entire year. The second goal was to take that code and then make changes so that it would run faster and could scale up to larger and larger datasets while still providing the same basic functionality.

## Analysis

I worked this project in an Excel workbook and wrote VBA subroutines to handle the analysis. We started with two tabs of data, "2017" and "2018", that had columns containing the stock ticker, trading date, closing price, and daily volume, among others. Thankfully the data were delivered in a clean, workable format and we didn't have to spend any time cleaning the spreadsheet up before getting to our code. Our code went through two iterations, the first one that got the job done and then a second one that did the same job but moved almost ten times quicker (making it a better candidate to use moving forward with larger datasets).

### First iteration code
 
 In the first iteration of our code, we took the path of looping through our ticker names and then, for each ticker symbol, looping through the entire data set to pick out that ticker's starting and ending prices along with its total volume. The core loops are shown below.
    
    For i = 0 To 11
        'clear out each variable for that ticker
        ticker = tickers(i)
        totalVolume = 0
                
        Worksheets(yearValue).Activate
        For j = rowStart To rowEnd
            
            'Find the first and last instances of the ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
            
            'check if stock is of the correct ticker, then increase totalVolume
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
        Next j

This completed our task, but we realized here that we were reading every row in the data tab 12 times and it would probably be much more efficient to run through the list once and gather all the information we need in that go.    

### Second iteration code
For this refactoring we decided to try and loop through all the data once and then store our information in separate arrays containing the total volume, starting price, and ending price for each ticker. The core loop is shown below.

    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'Increase the tickerIndex
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

## Refactoring Results

The results here showed that, as expected, our second iteration code was much more efficient than our first pass. The first iteration ran in around 0.5 seconds on my machine. Which obviously isn't a long time, but this is a very small data set and that could get unwieldy if we try to add larger swaths of the stock market to our analysis. Our second iteration ran in about 0.05 seconds on my machine (screenshots below). Just making the decision to learn about arrays and loop through the data only once made the code overall run about ten times faster. In the end it was worth the effort.

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

## Stock Results

As interesting as the code is, what Steve (and his parents) care about is the end results: how did these stocks do? In short, it looks like 2018 was a difficult year for green energy companies. Looking at all twelve companies together, there was an average return of 67.3% in 2017 and that dipped all the way down to -8.5% in 2018. 

While it's not always wise to use past performance as a predictor for future gains, it certainly seems worth looking more into ENPH and RUN as investment opportunities. ENPH had a return of 129.5% in 2017 and managed to avoid the worst of the dip in 2018 to still deliver an 81.9% return. And while RUN only had a 5.5% return in 2017, it managed an 84% return in 2018. In both cases it would seem prudent for Steve to look in to how these companies managed to deliver these returns in a year when comparable companies suffered and decide if he sees them as sustainable practices that would help his parents' portfolio by continuing those gains.

## Summary

- What are the advantages or disadvantages of refactoring code?

An advantage of refactoring code can be that it gives you increased efficiency, but for me what I found most useful was the practice of taking the same problem and trying to rethink how I solved it. That mental flexibility seems like it's only going to be a plus moving forward into more coding challenges. A disadvantage can be that it may take more time to make thoughtful, efficient code than it is to slap down some code that answers the question before moving on to your next problem. Depending on the situation you may prefer to the quick and dirty option and then save the nice answer for when you have a little more time to revisit the problem at hand.

- How do these pros and cons apply to refactoring the original VBA script?

The pros, as seen above, was the time save in running the code. The con here, not that it was a major one, was that I had to brush up on using arrays in VBA. I'm familiar with the data structure but can never remember when moving between languages if I use () or [] to access individual values, or if indices start at 0 or 1. Of course this added a little bit of time to the project but it was worth it in the end, so it wasn't a major con. 
