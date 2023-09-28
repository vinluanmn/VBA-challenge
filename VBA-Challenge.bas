Attribute VB_Name = "Module1"
Sub Worksheetloop():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    StockChange ws
Next ws

End Sub

Sub StockChange(ws As Worksheet):
'Define headers
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2:O4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    ws.Columns("I:Q").AutoFit
'Define variables
    Dim ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    Dim summaryTableRows As Integer
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    'Intializing variables
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    greatestIncreaseTicker = ""
    greatestDecreaseTicker = ""
    greatestVolumeTicker = ""
    
    
    'end of worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summaryTableRows = 2 'start at row 2
    
        'loop starting from row 2 until the end of the sheet (lastrow)
        For Row = 2 To lastrow
        'track changes in the first column
            If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
                OpeningPrice = ws.Cells(Row, 3).Value
                'ticker changes, set the ticker name variable
                ticker = ws.Cells(Row, 1).Value
                ws.Cells(summaryTableRows, 9).Value = ticker
            ElseIf ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ClosingPrice = ws.Cells(Row, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                ws.Cells(summaryTableRows, 10).Value = YearlyChange
                PercentChange = (YearlyChange) / (OpeningPrice)
                ws.Cells(summaryTableRows, 11).Value = PercentChange
                ws.Cells(summaryTableRows, 11).NumberFormat = "#0.00%"
                
                Dim changeCell As Range
                Set changeCell = ws.Cells(summaryTableRows, 10)
                If YearlyChange > 0 Then
                    changeCell.Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    changeCell.Interior.Color = RGB(255, 0, 0)
                End If
                
                If PercentChange > greatestIncrease Then
                    greatestIncrease = PercentChange
                    greatestIncreaseTicker = ticker
                ElseIf PercentChange < greatestDecrease Then
                    greatestDecrease = PercentChange
                    greatestDecreaseTicker = ticker
                End If
                
                StockVolume = StockVolume + ws.Cells(Row, 7).Value
                ws.Cells(summaryTableRows, 12).Value = StockVolume
                
                If StockVolume > greatestVolume Then
                    greatestVolume = StockVolume
                    greatestVolumeTicker = ticker
                End If
                
                'add one to the summary table rows (moves to the next row in the summary table)
                summaryTableRows = summaryTableRows + 1
                'reset openprice, closingprice and yearly change
                OpeningPrice = 0
                ClosingPrice = 0
                YearlyChange = 0
                PercentChange = 0
                StockVolume = 0
            Else
                StockVolume = StockVolume + ws.Cells(Row, 7).Value
            End If
        Next Row
                ws.Range("P2").Value = greatestIncreaseTicker
                ws.Range("Q2").Value = greatestIncrease
                ws.Range("Q2").NumberFormat = "#0.00%"
                ws.Range("P3").Value = greatestDecreaseTicker
                ws.Range("Q3").Value = greatestDecrease
                ws.Range("Q3").NumberFormat = "#0.00%"
                ws.Range("P4").Value = greatestVolumeTicker
                ws.Range("Q4").Value = greatestVolume
End Sub
