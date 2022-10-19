# -Stock-analysis-mac-
Stock analysis to assist Steve and his parents

# Overview of Project

The goal for this project was to assist Steve find the results of a given set of stocks over two years. Steve is wanting to assist his parents with making an educated choice of which stocks would have a higher propbability of generating positive returns. 
The attempt to refactor our initial code was to assist Steve in the future if he decides to analayze larger data sets at the same time. By refactoring the code this should help the macros run more efficiently in the future.

# Results

After refactoring the code the program will run at roughly 7 milliseconds where is the opriginal ran at roughly 13 milliseconds. This is roughly twice as fast. The reasons that this occurred is we simplified the macros in order for it to loop through the data using one formula rather than multiple, this will be useful moving in multiple applications as in this day and age speed is key and time is money.

<img width="347" alt="Screen Shot 2022-10-19 at 2 19 56 PM" src="https://user-images.githubusercontent.com/114188120/196778662-2c1eae3e-ded5-4009-89d6-b9fe2da2fca4.png">
<img width="258" alt="Screen Shot 2022-10-19 at 2 22 59 PM" src="https://user-images.githubusercontent.com/114188120/196778706-66aac110-d0d3-47b3-83c4-01e01a0ef1aa.png">

## Here is the refactored code 
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

'1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i

''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub

Sub formatAllStocksAnalysisTable()

    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Font.Size = 14
    
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

    
    
End Sub

Sub ClearWorksheet()

        Cells.Clear
        
End Sub


# Summary

As stated above, refactoring the code will allow for the program to run more effeciently, it also helps to keep the code more organized as there are fewer lines to accomplish the same end objective, this will also help in getting a quicker read or overview of the gooal should anyone need to reveiw the code at a later date. This can also be a disadvantage in the fact that you are having to go back in a reformat something that may already be functional and does not present each line item out if you are walking through step by step.

Any time you do go back in to refactor a code, it involves changing the original code. This obviously means that there will more than likely be trial and error periods, research periods, and of course potential frustration. There is always the possibility of human error when doing this and having a few areas mess up along the way. So the best bet is if you are on a time table for a project to do what works best in order to have the program run correctly, then return to it to improve upon what is already written if there was not an immediate way to run it more effeciently in the first place.
