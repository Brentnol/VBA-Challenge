<!-- Strong-->

# VBA Challenge Analysis
<!-- Horizontal Rule-->

## Overview and Background of Project
The Client (Steve) is quite comfortable with excel; however, he requires a more robust method to analyze stock data. Steve would like to find the total daily volume and yearly return for each stock for 2017 and 2018. The daily stock volume refers to the number of shares traded throughout the day and measures how actively a stock is traded. The yearly return usually represents this variable as a percentage difference in price from the beginning to the end of the year.      

## Purpose
The purpose of this Data Analytic exercise is to assist Steve, and by extension, Steve's parents, with a detailed analysis of a large excel stock dataset comprising a series of stocks for the years 2017 and 2018, with opening/ closing numbers and corresponding volumes, to determine the financial viability of stocks traded in this period. This will allow Steve and his parents to make an informed decision on which stock to pursue.

## Analysis and Challenges
The challenge of this exercise is to edit and refactor previously written codes to loop through all the data one time to collect the same information as the previous source code and to determine whether refactoring the original codes successfully made the VBA macro run faster, and subsequently, to produce an analysis that explains the findings thus:

### Initial Review of Dataset:
<!--UL-->
* First, a high-level review was done to get a sense of the data regarding the dataset's number of columns and rows. This was done to understand the data types and determine whether the data was readable or would need to be converted before progressing further with the analysis. At the same time, the file was saved as a Microsoft Excel macros-enabled worksheet to allow VBA scripting to progress in the back-end. 

<!--Links-->
<!--UL-->
* Because the Stock dataset was so large, an investigation was done to precisely establish the number of columns and rows. This was done by utilizing various keyboard shortcuts, combining keyboard keys such as CTRL/right or down arrow keys to get a quick idea of the dataset's number of rows and columns. It was established that Steve's data is stored in rows 1 through 3013 and in columns A through H in both year 2017 and 2018 datasets. 
  
* Next, it was necessary to determine whether VBA was installed correctly on the machine and permit VBA scripting to progress. This was achieved by building a simple subroutine to check macros. This was proven to be favourable.  
  
<!--UL-->
### Initial analysis to test data before deep dive:

* This was achieved by utilizing the Range() method, which selects cells with the same range format that Excell formulas use. This method also allows one to choose a range of only one cell. Hence, as a preliminary test, cell A1 was set to the value "DAQO (Ticker:DQ)" using the code Range("A1").value="DAQO (Ticker:DQ)" as follows:"

Sub DQAnalysis()
    
    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker: DQ)"

End Sub

This was followed by the use of the cell method thus:

Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
End Sub
 
* The ensure that the codes were readable, the writer used comments to explain components of the code to promote ease of use by others and assist the writer should Steve require future assistance to add/ expand or refactor the capability of the Stock Data code.  

 <!--UL--> 
## Analysis of All Stocks & Results 

* A new worksheet was created and named "All Stocks Analysis". This was used as the output for the analysis of multiple stocks options. 

* Next, a new subroutine was created called "AllStocksAnalysis" to format the output worksheet thus:

1. The "All Stock Analysis" worksheet was activated, and Cell A1 was coded to the title "All Stocks (year)". Three columns were added with headers; Ticker, Total Daily Volume and Return thus:

 'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

2. Followed by the initialized array of all the tickers:

'Initialize array of all tickers
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
<!--Links-->
3. Next,  [Stockoverflow](https://stackoverflow.com/questions/18088729/row-count-where-data-exists
"Row Count where data exist") was consulted for assistance with a code string to get the number of rows to loop over. The following line of code was used thus:  RowCount = Cells(Rows.Count, "A").End(xlUp).Row. The remainder of the code used to carry out the analysis were as follows: 
   
 '1a) Create a ticker Index
    Dim tickerIndex As String
    tickerIndex = 0
   
    '1b) Initialize variables for starting price and ending price
    Dim tickerVolume As Long
    Dim startingPrice As Single
    Dim endingPrice As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
   
       For i = 0 To 11
        ticker = tickers(i)
        tickerVolumes = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
     Worksheets(yearValue).Activate
        For j = 2 To RowCount
       
       '3a) Increase volume for current ticker
        
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
       '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
        
        End If
       '3c) check if the current row is the last row with the selected ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            endingPrice = Cells(j, 6).Value
        
        End If
        
    Next j
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

          
    Next i
    
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
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If

    Next i
      
   endTime = Timer
   MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year" & (yearValue)
End Sub
 
### Analysis of Outcomes Based Refactored Codes versus Original Codes
<!--UL-->
 <!--Strong--> 
 <!--Images-->
### Results from Analysis - Refactored Codes:

![VBA_Challenge 2018](https://github.com/Brentnol/VBA-Challenge/blob/main/Resources/VBA_Challenge_2018.png)

 
![VBA_Challenge 2017](https://github.com/Brentnol/VBA-Challenge/blob/main/Resources/VBA_Challenge_2017.png)

 
 
 
 <!--Strong-->   
### Results from Analysis - Original Codes:

![VBA_Challenge 2018](https://github.com/Brentnol/VBA-Challenge/blob/main/Resources/VBA_Challenge_2018_Original.png)

![VBA_Challenge 2017](https://github.com/Brentnol/VBA-Challenge/blob/main/Resources/VBA_Challenge_2017_Original.png)

 <!--OL--> 
### Challenges and Solutions Encountered

1.	Learning and executing the for loop, Nested loop, and array commands were challenging. However, over time, I grasped key concepts of these commands to produce the desired results.

2.	Stockoverflow was a fantastic resource with a considerable repertoire of codes and other valuable resources. StackOverflow was consulted extensively during this exercise. 

## Conclusions:
<!--OL-->
### Outcomes based on Refactored vs Original Codes:
1.	The initial code was refactored using the starter code for this exercise and was shown to reduce the run time from 1.25 seconds to 1.19 seconds for the 2018 stocks and from 1.46 seconds to 1.21 seconds for the 2017 stocks. This supports the argument that code refactoring helps to speed up the runtime of macros while simplifying support and code updates. At the same time, ensuring clean codes are written and produced with reduced complexity and ease of understanding. Furthermore, this helps with the maintainability and scalability of codes.     
