# stock-analysis
## Overview of Project

- In the Stock Analysis project, we will help Steve with the analysis of a green stock for his parents to see if it is worth investing in. Steve's parents are passionate about green energy. So, they decided to invest in DAQO New Energy Corp. DAQO's ticker symbol is "DQ". So further we'll be using the ticker symbol "DQ" in our analysis.

### DQ Analysis

- As Steve's parents are starting to pester him about DAQO's stock we'll be starting our analysis with "DQ" .
-  We are using VBA code to further help Steve in stock analysis.
- First we need to started with creating Macros and write the code under that particular macro which was DQ Analysis Shown as below :

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    cells(3, 1).Value = "Year"
    cells(3, 2).Value = "Total Daily Volume"
    cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + cells(i, 8).Value

        End If

        If cells(i - 1, 1).Value <> "DQ" And cells(i, 1).Value = "DQ" Then

            startingPrice = cells(i, 6).Value

        End If

        If cells(i + 1, 1).Value <> "DQ" And cells(i, 1).Value = "DQ" Then

            endingPrice = cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    cells(4, 1).Value = 2018
    cells(4, 2).Value = totalVolume
    cells(4, 3).Value = (endingPrice / startingPrice) - 1
End Sub

- As per above code  and after analysis the DQ stock for Steve and his parents and came to conclusion that the DQ does not have good returns in 2018.

### All Stock Analysis

- So as per above analysis steve and  his parents wants to analyzing multiple stocks options to find out good return for them
- We define new sub as All stocks analysis for further coding that shows below:

Sub AllStocksAnalysis()

   '1) Format the output sheet on All Stocks Analysis worksheet
    Dim startTime As Single
    Dim endTime  As Single

  yearValue = InputBox("What year would you like to run the analysis on?")

  startTime = Timer
   Worksheets("All Stocks Analysis").Activate
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   cells(3, 1).Value = "Ticker"
   cells(3, 2).Value = "Total Daily Volume"
   cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   
   Dim tickers(11) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   
   '3a) Initialize variables for starting price and ending price
   
   Dim startingPrice As Long
   Dim endingPrice As Long
   
   '3b) Activate data worksheet
   
   Worksheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   
   RowCount = cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   
     For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
       
       '5a) Get total volume for current ticker
       
           If cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + cells(j, 8).Value

           End If
           
        '5b) get starting price for current ticker
        
           If cells(j - 1, 1).Value <> ticker And cells(j, 1).Value = ticker Then

               startingPrice = cells(j, 6).Value

           End If

        '5c) get ending price for current ticker
        
           If cells(j + 1, 1).Value <> ticker And cells(j, 1).Value = ticker Then

               endingPrice = cells(j, 6).Value

           End If
           
       Next j
       
        '6) Output data for current ticker
        
           Worksheets("All Stocks Analysis").Activate
           
            cells(4 + i, 1).Value = ticker
            cells(4 + i, 2).Value = totalVolume
            cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   endTime = Timer
   
   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If cells(i, 3) > 0 Then
            
         cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 End Sub
- We will go through the analysis from the code and conclude the Results.

### Results

- First Lets see the result of DQ stocks , We calculate the total daily volume and yearly return, as per below result 
Daqo might not be the best option for Steve's parents to invest in.

![Screenshot (37)](https://user-images.githubusercontent.com/96400887/166111857-6dc18ef8-de69-43e8-80f7-8fa722b91bee.png)

- Then we did All stock Analysis for 2018 year and results shows as below:

![2018_Refactor code data](https://user-images.githubusercontent.com/96400887/166121310-31f1c86b-9657-4883-965e-fec65f3827b6.png)

- In the future, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results. To help Steve, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box
- With the refactoring code we do analysis and help steve and his parents to understand what to invest in.
-Refactoring is a key part of the coding process. we are taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read, below is the refactor code:

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    cells(3, 1).Value = "Ticker"
    cells(3, 2).Value = "Total Daily Volume"
    cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
       
       tickerIndex = 0
    

    '1b) Create three output arrays
    
       Dim tickerVolumes(12) As Long
       Dim tickerStartingPrices(12) As Single
       Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
       For i = 0 To 11
       
        tickerVolumes(i) = 0
        
       Next i
        
        'Activate data worksheet
       Worksheets(yearValue).Activate
        
    ''2b) Loop over all the rows in the spreadsheet.
    
        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
         For tickerIndex = 0 To 11
         
          If cells(j, 1).Value = tickers(tickerIndex) Then
          
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + cells(j, 8).Value
          
          End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If cells(j - 1, 1).Value <> tickers(tickerIndex) And cells(j, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = cells(j, 6).Value
            
            
           End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
            
            If cells(j + 1, 1).Value <> tickers(tickerIndex) And cells(j, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = cells(j, 6).Value
            
            End If
       

        '3d Increase the tickerIndex.
        
         Next tickerIndex
             
        Next j
    
   
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
         Worksheets("All Stocks Analysis").Activate
        For i = 0 To 11
        
         cells(4 + i, 1).Value = tickers(i)
         cells(4 + i, 2).Value = tickerVolumes(i)
         cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
         
     Next i
     
      Worksheets("All Stocks Analysis").Activate

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If cells(i, 3) > 0 Then
            
         cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    
    End Sub
    

-  Steve will probably want to run this analysis for each year, so we have analyzed data of 2017 and 2018 with refactored code as shown below:

![2017_VBA Challenge](https://user-images.githubusercontent.com/96400887/166122587-a240c497-cc45-4d61-9650-1acf77d6a85f.png)

![2018_VBA Challenge](https://user-images.githubusercontent.com/96400887/166122594-79d44c30-1d62-47a2-9efe-7d0bffc2bdef.png)

- To see the result we need to compare refactored code vs original code , Lets see the original code runtime for both the years as shown below :

![2017_original code](https://user-images.githubusercontent.com/96400887/166123277-2dc66a63-4d32-4bfe-8d37-dff01669d9eb.png)

![2018_original code](https://user-images.githubusercontent.com/96400887/166123286-81b4cc2e-7ea5-4750-afc4-263f0eebb7c8.png)

- **We can finally conclude that refactoring code screen running time is less then the original script for both the years and the refactoring is more easier to understand and read.**

### Summary

- Steve may want to look at a different set of stocks in the future. With this in mind, we created a flexible macro for running multiple stocks. By carefully reusing the code we've already written for DQ, we wrote a macro with this flexibility
- To run analyses on all of the stocks, we created a program flow that loops through all of the tickers.
- Running an analysis of all stocks—is to copy the code from the Daqo analysis and paste it over and over, changing the ticker and the line to output each time.  
- we've run the analysis, to make it easier for Steve to read by adding some formatting to our table. Like changing font styles, adding borders, setting number formats, and so on—but we can automate formatting with VBA.
- Steve needs a way to run these analyses. He could install the Developer tab, but a button would be easier and more user-friendly so we made a button for Steve(Run Analysis for all stocks)
- Steve want to run this analysis for each year, so we updated our code to run for any year, not just 2018
- For more efficient and faster result we refactored our code.

1. ### What are the advantages or disadvantages of refactoring code?

-    The advantage of refactoring code is the screen running time is less then the original script and the refactoring is more easier to understand and read.
- But the only disadvantage i realized is it is time consuming and very complex to "write" the script
As we saw earlier the refactored code is more tedious.

2. ### How do these pros and cons apply to refactoring the original VBA script?

- I have noticed that the resources were utilized less as compared to the original vba script and cons is the refactoring is more complex structure.
- When we calculated the same total daily volume and returns for 2017 and 2018 years refactored code vs original VBA script, it shows refactored code took comparatively less time which we noticed in the screenshot earlier, and below are the results for better understanding,

     Run time for refactored code 2017 (0.644453 seconds per year) VS 
     Run time for original VBA script 2017 (0.707031 seconds per year)

     Run time for refactored code 2018 (0.625 seconds per year) VS 
     Run time for original VBA script 2018 (0.699 seconds per year).



