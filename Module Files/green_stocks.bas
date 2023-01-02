Attribute VB_Name = "Module1"
Sub MacroCheck()

    Dim testMessage As String

    testMessage = "Hello World!"

    MsgBox (testMessage)

End Sub

Sub DQAnalysis()
    Worksheets("DQ_Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ_Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1

End Sub
Sub clearDQAnalysis()
    
    'Activate All Stocks Analysis worksheet
    Worksheets("DQ_Analysis").Activate

    Rows(4).clearContents
    
    Columns("C").Interior.Color = xlNone

End Sub

Sub formatAllStocksAnalysisTable()

    Dim startTime As Single
    Dim endTime  As Single
    
    startTime = Timer

   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (2018)"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

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
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets("2018").Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i



    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.0# "
    Range("C4:C15").NumberFormat = "0.0#%"
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
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub

Sub clearAllStocksAnalysis()
    
    'Activate All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Rows(4).clearContents
    Rows(5).clearContents
    Rows(6).clearContents
    Rows(7).clearContents
    Rows(8).clearContents
    Rows(9).clearContents
    Rows(10).clearContents
    Rows(11).clearContents
    Rows(12).clearContents
    Rows(13).clearContents
    Rows(14).clearContents
    Rows(15).clearContents
    
    Columns("C").Interior.Color = xlNone

End Sub

Sub yearValueAnalysis()
    
    Dim startTime As Single
    Dim endTime  As Single
    
    'Current issue with this code is that an error is thrown if the user
    'inputs an incorrect date - requiring the code to be deleted and
    'restarted in order to reset the input
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("Year Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    Dim startingPrice As Single
    Dim endingPrice As Single
   
    '3b) Activate data worksheet
    Worksheets(yearValue).Activate
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        '5) loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value

            End If
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If
        Next j
       
        '6) Output data for current ticker
        Worksheets("Year Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i



    'Formatting
    Worksheets("Year Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.0# "
    Range("C4:C15").NumberFormat = "0.0#%"
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
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub

Sub ClearYearAnalysis()
    
    'Activate All Stocks Analysis worksheet
    Worksheets("Year Analysis").Activate

    Rows(4).clearContents
    Rows(5).clearContents
    Rows(6).clearContents
    Rows(7).clearContents
    Rows(8).clearContents
    Rows(9).clearContents
    Rows(10).clearContents
    Rows(11).clearContents
    Rows(12).clearContents
    Rows(13).clearContents
    Rows(14).clearContents
    Rows(15).clearContents
    
    Columns("C").Interior.Color = xlNone

End Sub
