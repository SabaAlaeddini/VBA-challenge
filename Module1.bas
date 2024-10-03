Attribute VB_Name = "Module1"
Sub quarter_pt1()

    ' Create variables
    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker_Symbol As String
    Dim Table_Row As Long
    Dim Q_Open_Price As Double
    Dim Q_Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Counter As Long
    Dim Total_Stock_Volume As Double
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim MaxTotalValue As Double
    Dim MaxValueRow As Long
    Dim MinValueRow As Long
    Dim MaxTotalValueRow As Long
    Table_Row = 2
    Counter = 0
    Total_Stock_Volume = 0    ' LOOP THROUGH ALL SHEETS
    
    
    For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decreast"
        ws.Range("O4").Value = "Greatest Total Volume"        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then                ' Set the Ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value                ' Set the Opening Price
                Q_Open_Price = ws.Cells(i - Counter, 3).Value                ' Set the Closing Price
                Q_Close_Price = ws.Cells(i, 6).Value                ' Calculate the Quarterly Change
                Quarterly_Change = Q_Close_Price - Q_Open_Price                ' Calculate the TOtal Stuck Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value                ' Print the Ticker symbol in each Table
                ws.Range("I" & Table_Row).Value = Ticker_Symbol                ' Print the Quarterly Change in each Table
                ws.Range("J" & Table_Row).Value = Quarterly_Change
                ws.Range("J" & Table_Row).NumberFormat = "0.00"
                
                If Quarterly_Change > 0 Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                ElseIf Quarterly_Change < 0 Then
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Table_Row).Interior.ColorIndex = 0
                End If                ' Print the Percent Change in each Table
                
                
                If Q_Open_Price <> 0 Then
                    ws.Range("K" & Table_Row).Value = (Quarterly_Change / Q_Open_Price)
                    ws.Range("K" & Table_Row).NumberFormat = "0.00%"
                    Else
                    ws.Range("K" & Table_Row).Value = 0
                End If                ' Print the Total Stock Volume in each Table
                
                
                ws.Range("L" & Table_Row).Value = Total_Stock_Volume                ' Add one to the table row
                Table_Row = Table_Row + 1                ' Reset the Counter
                Counter = 0                ' Reset the Total Stock Volume
                Total_Stock_Volume = 0
                Else
                Counter = Counter + 1
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If
            Next i        ' Reset table row for next sheet
            
            
        Table_Row = 2        ' Calculate the Maximum Value in column K
        MaxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))        ' Find the row of the Maximum Value using a loop
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = MaxValue Then
                MaxValueRow = i
                Exit For
            End If
        Next i        ' Print the ticker and value in columns P and Q
        
        
        ws.Range("P2").Value = ws.Range("I" & MaxValueRow).Value
        ws.Range("Q2").Value = MaxValue
        ws.Range("Q2").NumberFormat = "0.00%"        ' Calculate the Minimum Value in column K
        MinValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))        ' Find the row of the Minimum Value using a loop
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = MinValue Then
                MinValueRow = i
                Exit For
            End If
        Next i
        
        
        ws.Range("P3").Value = ws.Range("I" & MinValueRow).Value
        ws.Range("Q3").Value = MinValue
        ws.Range("Q3").NumberFormat = "0.00%"        ' Calculate the Maximum Total Stock Volume in column L
        MaxTotalValue = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))        ' Find the row of the Maximum Total Value using a loop
        
        For i = 2 To LastRow
            If ws.Cells(i, 12).Value = MaxTotalValue Then
                MaxTotalValueRow = i
                Exit For
            End If
        Next i
        
        ws.Range("P4").Value = ws.Range("I" & MaxTotalValueRow).Value
        ws.Range("Q4").Value = MaxTotalValue        ' Autofit to display data
        ws.Columns("A:O").AutoFit
        
          Next ws
        End Sub

