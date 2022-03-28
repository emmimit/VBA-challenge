'VBA Homework - The VBA of Wall Street
Sub StockMarketAnalysis()



'PART I

Dim ws As Worksheet
Dim i As Long
Dim RowCount As Long
Dim Ticker As String
Dim Volume As Double



    'Create a loop for each worksheet in the workbook
For Each ws In Worksheets


    'Create header columns for each worksheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Volume"


    'Find the last row of each worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Volume = 0

Dim Summary_table_row As Integer
Summary_table_row = 2


    'Looping through the list of tickers (Skipping the header row)
For i = 2 To LastRow

    
    'Check if we are still within the same Ticker Name, If we are not
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value
    
    ws.Range("I" & Summary_table_row).Value = Ticker
    ws.Range("J" & Summary_table_row).Value = Volume
    
    Volume = 0
    
    
    Summary_table_row = Summary_table_row + 1
    
    Else
     Volume = Volume + ws.Cells(i, 7).Value
     
    End If
    
    Next i
    
    
    'Store data value in variables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    'Define variables
    Dim YearlyChange As Double
    Dim PercentCHange As Double
    
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Volume = 0
    
    Summary_table_row = 2
    
    
    For i = 2 To LastRow
    
    'Calculate stock volume
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value
    
     ws.Range("I" & Summary_table_row).Value = Ticker
     ws.Range("L" & Summary_table_row).Value = Volume
     
     Volume = 0
     
     'Closing price for the ticker at the end of the year
     ClosingPrice = ws.Cells(i, 6)
     
     'Get yearly change value and add the value to relevant cell in each of the worksheets
     If OpeningPrice = 0 Then
        YearlyChange = 0
        PercentCHange = 0
    Else:
        YearlyChange = ClosingPrice - OpeningPrice
        PercentCHange = (ClosingPrice - OpeningPrice) / OpeningPrice
    End If
        
        ws.Range("J" & Summary_table_row).Value = YearlyChange
        ws.Range("K" & Summary_table_row).Value = PercentCHange
        ws.Range("K" & Summary_table_row).Style = "Percent"
        ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
    
      Summary_table_row = Summary_table_row + 1
        
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    
        OpeningPrice = ws.Cells(i, 3)
        
    Else:
    Volume = Volume + ws.Cells(i, 7).Value
    
     End If
     
    
     Next i
     
    For i = 2 To LastRow
    
    'Applying conditional formatting by highlighting positive and negative yearly change values.
     If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4
     
     ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
     
     
     End If
     
     Next i
     
     
     'PART II : CHALLENGES
     
     'Create Challenge Table to show, greatest % Increase, decrease and greatest total volume
     
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     
     Dim GreatestIncrease As Double
     Dim GreastestDecrease As Double
     Dim GreatestVolume As Double
     
     GreatestIncrease = 0
     GreatestDecrease = 0
     GreatestVolume = 0
     
     For i = 2 To LastRow
     
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("Q2").Style = "Percent"
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("P2").Value = ws.Cells(i, 9).Value
        End If
        
     Next i
     
    For i = 2 To LastRow
    
        If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("Q3").Style = "Percent"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("P3").Value = ws.Cells(i, 9).Value
        End If
        
    Next i
    
   For i = 2 To LastRow
   
        If ws.Cells(i, 12).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(i, 12).Value
            ws.Range("Q4").Value = GreatestVolume
            ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
        
    Next i
    
 ws.Columns("A:Q").AutoFit
 

 Next ws
        

End Sub