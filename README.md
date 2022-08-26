# Stock Analysis Project

In this project, I analyzed green energy stocks for Steve’s parents who are interested in investing in green energy but haven’t done a lot of research yet. Steve wanted to analyze a few *green energy stocks* besides the *Daqo stock* that his parents had already invested in. 

Initially, I calculated the **Daily Volume** and **Yearly Return** for the *Daqo stock* using an Excel macro to run the analysis. One of the findings was that the Daqo stock had dropped **62.6%** in **2018**. Since Daqo wasn’t the best option for Steve’s parents, I created a macro to analyze all the stocks to find a better option. 
I reused part of the first code to do an **All Stocks Analysis** which simplified the task. 

Subsequently, Steve asked to broaden the analysis to the stock market. One of the challenges with this task is that this analysis will have a large data set, and the code I have worked with so far will unlikely perform as well. 

## VBA Challenge Results

Now that I was going to work with a large data set I decided to refactor the **All Stocks Analysis** code to improve time performance. The following code is the first code I used to analyze the DQ stock. 
 
```    'loop over all the rows
  For i = rowStart To rowEnd
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
        'set starting price
            startingPrice = Cells(i, 6).Value
        End If
            
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
        'set ending price
            endingPrice = Cells(i, 6).Value
        End If
        
  Next i 
   ```  
     

Since **DQ analysis** only required to go through one *stock* or *ticker*, I used only one ```for loop```. However, to go through all the stocks I modified the code to include a nested ```for loop```, as well as an ```inputBox``` to allow the user to choose the year they wanted to perform the analysis on.

The following is a sample of the modified code for **All Stock Analysis**.

```
'Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
   
        'Worksheets("2018").Activate changed to
        Sheets(yearValue).Activate
       'Loop through rows in the data.
        For j = 2 To RowCount
            'Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            End If

            'Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            End If
            
            'Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            End If
            
        Next j
   'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
        
```

I added a ```Timer``` function  to see how long it took the code to output the results. For year 2017 it took ≈ 0.3945 seconds, and for year 2018 it took ≈0.3867. The following images are the message boxes showing the elapsed time.

![image_name](/Ressources/Original_Code_2017.png)

![image_name](/Ressources/Original_Code_2018.png)


Considering that I was going to work with a larger dataset, this is how I refactored the **All Stocks Analysis** code for the ***VBA Challenge***

```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Calculate volumes for each ticker 
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'Intialized a ticker Index to loop through tickers array
    tickerIndex = 0
    
    'Loop over all the rows in the spreadsheet
     For i = 2 To RowCount
            'sum tickerVolume only if current cell ticker equals the value of the selected ticker
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
    
            'find starting price by cheking if the previous ticker does not equal the selected ticker
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
            'if next ticker does not equal the selected ticker set ending price and increase ticker index
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
            End If
    
     Next i
    
    'This for loop goes through the arrays in order to output the Ticker, Total Daily Volume, and Return.
     For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
     Next i
 ```
We can observe some differences between the last two codes. In the **All Stock Analysis** code we have one nested ```for loop```, and in the **VBA Challenge** code we have three separate ```for loops``` and four different *arrays*.

When running this subroutine, I found that the elapsed time decreased significantly. In other words it ran a lot faster! As you can see below for year 2017 it only took ≈ 0.0644 seconds, and for year 2018 it took ≈0.0625.

![image_name](/Ressources/VBA_Challenge_2017.png)

![image_name](/Ressources/VBA_Challenge_2018.png)

Since I noticed that the refactored code did not use nested loops, I googled nested loops vs non nested loops performance and found some interesting answers. In [Quora](https://www.quora.com/Which-is-better-a-nested-loop-with-particular-depth-or-the-same-number-of-loops-one-after-the-other) I found that most people agreed that non-nested loops are more efficient due to the difference in time complexity for nested loops and non-nested loops respectively. People's answers in Quora go into very technical and mathematical explanations, but I was able to verify that the refactored code without nested loops executed much faster than the original code with nested loops. The refactored code ran in about 60 miliseconds, and it took almost half a second to run to the original code. 



### Advantages and disadvantages

Found in [Stackoverflow](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software).
