Sub StockAnalysis()

For Each ws In Worksheets

'Set variables for Stock Analysis VBA Challenge
Dim Ticker As String
Dim LastRow As Long

'Variable for Total Stock Volume for the Year - used As String for Volume data, original tried As Long but it caused an Overflow error
Dim TotStockVol As String
TotStockVol = 0

Dim StockSumTable As Integer
StockSumTable = 2

Dim OpenPrice As Double
Dim ClosePrice As Double

Dim YearlyChg As Double
Dim PercentChg As Double

Dim OpenPriceLine As Double
OpenPriceLine = 2
 


    'Display Column Heading for Stock Summary Table, Format Cells and AutoFit Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Open"
        ws.Range("K1").Value = "Close"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percentage Change"
        ws.Range("N1").Value = "Total Stock Number"
        ws.Range("I1:N1").Font.Bold = True
        ws.Range("L:N").EntireColumn.AutoFit
        ws.Columns(12).NumberFormat = "0.00"
        ws.Columns(13).NumberFormat = "0.00%"
        ws.Columns(10).Hidden = True
        ws.Columns(11).Hidden = True
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        ws.Range("R1:S1").Font.Bold = True
        ws.Range("Q1").EntireColumn.AutoFit
        ws.Cells(2, 19).NumberFormat = "0.00%"
        ws.Cells(3, 19).NumberFormat = "0.00%"
     


' Determine Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    'Loop Through Stock Data
        For i = 2 To LastRow

        'Sum Total Stock Volume for the Year
        TotStockVol = TotStockVol + ws.Cells(i, 7).Value
    
      'Check for same Ticker Symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
        'Get the Ticker Symbol
           Ticker = ws.Cells(i, 1).Value
                   
        'Display Ticker Symbol to the Stock Summary Table
           ws.Range("I" & StockSumTable).Value = Ticker
                     
        'Display Total Stock Volume for the Year by Ticker
            ws.Range("N" & StockSumTable).Value = TotStockVol
                      
        'Reset the Total Stock Volume
            TotStockVol = 0
        
        'Get the Open Price, Closing Price, Yearly Change and the Percentage Change
        'Note: Displayed Open Price and Close Price on the worksheets, but the columns are hidden
        
            OpenPrice = ws.Range("C" & OpenPriceLine)
            ClosePrice = ws.Range("F" & i)
            YearlyChg = ClosePrice - OpenPrice
            ws.Range("J" & StockSumTable).Value = OpenPrice
            ws.Range("K" & StockSumTable).Value = ClosePrice
            ws.Range("L" & StockSumTable).Value = YearlyChg
            If ws.Range("L" & StockSumTable).Value >= 0 Then
                ws.Range("L" & StockSumTable).Interior.ColorIndex = 4
            Else
                ws.Range("L" & StockSumTable).Interior.ColorIndex = 3
            End If
            
        
        If OpenPrice = 0 Then
            PercentChg = 0
        Else
            OpenPrice = ws.Range("C" & OpenPriceLine)
            PercentChg = YearlyChg / OpenPrice
            ws.Range("M" & StockSumTable).Value = PercentChg
        End If
        
        'Add Rows to the Stock Summary Table
        StockSumTable = StockSumTable + 1
        OpenPriceLine = i + 1
        
      End If
    
    Next i

'Analyses for Greatest % Increase, Greatest % Decrease and Greatest Volume
    LastRow = ws.Cells(Rows.Count, 13).End(xlUp).Row
    For i = 2 To LastRow
    
        If ws.Range("M" & i).Value > ws.Range("S2").Value Then
            ws.Range("S2").Value = ws.Range("M" & i).Value
            ws.Range("R2").Value = ws.Range("I" & i).Value
        End If
        
        If ws.Range("M" & i).Value < ws.Range("S3").Value Then
            ws.Range("S3").Value = ws.Range("M" & i).Value
            ws.Range("R3").Value = ws.Range("I" & i).Value
        End If
        
        If ws.Range("N" & i).Value > ws.Range("S4").Value Then
            ws.Range("S4").Value = ws.Range("N" & i).Value
            ws.Range("R4").Value = ws.Range("I" & i).Value
        End If

     Next i

Next ws

End Sub