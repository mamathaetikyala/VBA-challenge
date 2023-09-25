Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim Ticker_Total_Volume As Double
Dim ticker As String
Dim ticker_start_date As Double
Dim ticker_end_date As Double
Dim Ticker_Open_Price As Double
Dim Ticker_Close_Price As Double

'Looping inWorksheets
For Each ws In ThisWorkbook.Worksheets

    'Clear Contents in Target Location before start
    ws.Columns("I:Q").ClearContents

    'Filling Headers for Target Locations
    ws.Cells(1, "i").Value = "Ticker"
    ws.Cells(1, "j").Value = "Yearly Change"
    ws.Cells(1, "k").Value = "Pecent Change"
    ws.Cells(1, "l").Value = "Total Stock Volume"
    ws.Cells(1, "p").Value = "Ticker"
    ws.Cells(1, "q").Value = "Value"
    ws.Cells(2, "o").Value = "Greatest%increase"
    ws.Cells(3, "o").Value = "Greatest%decrease"
    ws.Cells(4, "o").Value = "Greatest total volume"
    
    
    'Sort Data by Ticker and Date ascending
    ws.Columns.Sort key1:=ws.Columns("A"), Order1:=xlAscending, key2:=ws.Columns("B"), Order2:=xlAscending, Header:=xlYes
                    
    'Format Date Column to Numbers
    With ws.Columns(2)
        .NumberFormat = "0"
        .Value = .Value
    End With
    
    'Get Start Date and End date of Year
    Set r = ws.Range("B2:B" & ws.Rows.Count)
    ticker_end_date = Application.WorksheetFunction.Max(r)
    ticker_start_date = Application.WorksheetFunction.Min(r)
    
    
    
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    I = 2 'Skip Headers in Target Location and Start filling data from Second Row
    
    
   'Looping in Each Row in Worksheets
   For Each rw In ws.Rows
   
   
    'Skip Headers from Raw Data
    If rw.Row > 1 Then
   
        'On Every New Ticker
        If ws.Cells(rw.Row - 1, "A").Value <> ws.Cells(rw.Row, "A").Value Then
            ticker = ws.Cells(rw.Row, "A").Value
            ws.Cells(I, "i").Value = ws.Cells(rw.Row, "A").Value
            Ticker_Total_Volume = 0 'Initialize Total Ticker Volume to Zero
            ticker_start_date = Format(ws.Cells(rw.Row, "B").Value, 0)
            j = I
            I = I + 1
        End If
        
        'Adding Volume for Ticker into a Variable
        Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(rw.Row, "G").Value
        ws.Cells(j, "L").Value = Ticker_Total_Volume
        
        'Get Ticker Open Price
        If ws.Cells(rw.Row, "A").Value = ticker And ws.Cells(rw.Row, "B").Value = ticker_start_date Then
         Ticker_Open_Price = ws.Cells(rw.Row, "C").Value
        End If
        
        'Get Ticker Close Price, Year Price Change, Percent Difference and Conditional format with color
        'Find Greatest Volume, Percent Increase and Decrese
        If ws.Cells(rw.Row, "A").Value = ticker And ws.Cells(rw.Row, "B").Value = ticker_end_date Then
         Ticker_Close_Price = ws.Cells(rw.Row, "F").Value
         ws.Cells(j, "J").Value = (Ticker_Close_Price - Ticker_Open_Price)
         ws.Cells(j, "K").Value = FormatPercent((Ticker_Close_Price - Ticker_Open_Price) / Ticker_Open_Price)
         
         'Conditional Format Negative Change with Red and Positive Change with Green
         If ws.Cells(j, "J").Value < 0 Then
            ws.Cells(j, "J").Interior.Color = vbRed
         ElseIf ws.Cells(j, "J").Value > 0 Then
            ws.Cells(j, "J").Interior.Color = vbGreen
         End If
         
         'Find Greatest total volume
         If ws.Cells(j, "L").Value > ws.Cells(4, "Q").Value Then
            ws.Cells(4, "P").Value = ticker
            ws.Cells(4, "Q").Value = ws.Cells(j, "L").Value
         End If
         
         
        'Find Greatest Percentage Increase
         If ws.Cells(j, "K").Value > ws.Cells(2, "Q").Value And ws.Cells(j, "K").Value > s0 Then
            ws.Cells(2, "P").Value = ticker
            ws.Cells(2, "Q").Value = FormatPercent(ws.Cells(j, "K").Value)
        
         End If
         
         'Find Greatest Percentage Decrease
         If ws.Cells(j, "K").Value < ws.Cells(3, "Q").Value And ws.Cells(j, "K").Value < s0 Then
            ws.Cells(3, "P").Value = ticker
            ws.Cells(3, "Q").Value = FormatPercent(ws.Cells(j, "K").Value)
        
         End If
        End If
        
   
    End If
    
    'When last row is encountered exit loop
    If rw.Row = LastRow Then
        Exit For
    End If
   
   Next rw

Next ws

End Sub

