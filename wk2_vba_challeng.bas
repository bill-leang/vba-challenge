Attribute VB_Name = "Module1"
' Summarize the yearly statistics for each stock in the table (arranged in chronological order by each stock)
' done: Make it run for every sheet at once
' done: Find the stocks with the largest increase, largest decrease, and largest volumnes
' and output to a table two col away from the summary table

Sub CalculateSummaryData()
  Dim ws As Worksheet

  ' Calculate for every worksheet
  For Each ws In ThisWorkbook.Sheets
   ' Check that the worksheet contains stock data, otherwise do nothing.
   If ws.Cells(1, 1) = "<ticker>" Then
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim YearlyPercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    
    ' Initialize the summary table
    SummaryRow = 2 ' First row for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    ' Find the last row number in the data, similar as Ctl + Up
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Initialize variables
    YearOpen = ws.Cells(2, 3).Value ' Initial opening price
    Ticker = ws.Cells(2, 1).Value ' Initial ticker symbol
    TotalVolume = 0

    ' Loop through the data
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> Ticker Then
            ' Ticker has changed
            YearClose = ws.Cells(i, 6).Value ' Closing price
            YearlyChange = YearClose - YearOpen
            If YearOpen <> 0 Then
                YearlyPercentChange = (YearClose - YearOpen) / YearOpen
            Else
                YearlyPercentChange = 0
            End If
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ' Output results to the summary table
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = YearlyPercentChange
            ws.Cells(SummaryRow, 12).Value = TotalVolume

            ' Move to the next row in the summary table
            SummaryRow = SummaryRow + 1

            ' Reset variables for the next ticker
            Ticker = ws.Cells(i + 1, 1).Value
            YearOpen = ws.Cells(i + 1, 3).Value
            TotalVolume = 0
        Else
            ' Accumulate volume for the same ticker
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
    Next i

    ' Format Yearly change cells in the summary table as percentage and two decimal places
    ws.Range("K2:K" & SummaryRow).NumberFormat = "0.00%"
    
    
    ' Add conditional formatting for Yearly Change and Percent Change
    With ws.Range("J2:K" & SummaryRow - 1)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive changes
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative changes
        .FormatConditions(1).StopIfTrue = False ' Stop if first condtion is met
    End With
    
    
    ' BONUS
    ' Find the stocks with the largest increase, largest decrease and largest vol
    ' loop through the summary table
    ' keep track of largest_incr, largest_decr, largest_vol, if found larger, take note of the row number
    ' populate the third table
    
    Dim max_incrRow As Long
    Dim max_decrRow As Long
    Dim max_volRow As Long
    Dim max_incr As Double
    Dim max_decr As Double
    Dim max_vol As LongLong
    
    max_incrRow = 0
    max_decrRow = 0
    max_volRow = 0
    max_incr = 0
    max_decr = 0
    max_vol = 0
    
    ' The last row of summary table
    lastSTRow = ws.Cells(1, 9).End(xlDown).Row
    
    For i = 2 To lastSTRow
      ' find the largest % incr, decr, and vol
      ' if found, update max_incr, max_decr, max_vol and note the row number
      If ws.Cells(i, 11).Value > max_incr Then
        max_incr = ws.Cells(i, 11).Value
        max_incrRow = i
      End If
      If ws.Cells(i, 11).Value < max_decr Then
        max_decr = ws.Cells(i, 11).Value
        max_decrRow = i
      End If
      If ws.Cells(i, 12).Value > max_vol Then
        max_vol = ws.Cells(i, 12).Value
        max_volRow = i
      End If
    Next i
    
    ' Create the summary table
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    ' Set tickers
    ws.Cells(2, 16) = ws.Cells(max_incrRow, 9)
    ws.Cells(3, 16) = ws.Cells(max_decrRow, 9)
    ws.Cells(4, 16) = ws.Cells(max_volRow, 9)
    
    ' Set the values
    ws.Cells(2, 17) = ws.Cells(max_incrRow, 11)
    ws.Cells(3, 17) = ws.Cells(max_decrRow, 11)
    ws.Cells(4, 17) = ws.Cells(max_volRow, 12)
    
    ' format as percentage
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ' adjust width of columns
    ws.Columns("I:Q").AutoFit
   
   End If

  Next ws
End Sub

