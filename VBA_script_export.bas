Attribute VB_Name = "Module1"
Sub Stocks():
    Dim Ticker As String
    Dim ws As Worksheet
    Dim TotalTicker As Variant
    Dim SummaryRow As Integer
    Dim Percentchange As Double
    Dim MaxVal As Variant
    Dim MinVal As Variant
    Dim HTV As Variant


    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets
        'set the stock in row 2
        SummaryRow = 2

        openprice = ws.Cells(2, 3).Value
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'loop into rows 2 to lastrow, and i represents 2 to lastrow
        For i = 2 To lastrow
            Column = 1

            ' BEGIN IF STATEMENT that is triggered when we encounter a new ticker
            If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
                Ticker = ws.Cells(i, 1).Value
                'put the ticker symbol in the summary table
                ws.Cells(SummaryRow, 8).Value = Ticker
                ' get the value of the closing price
                closeprice = ws.Cells(i, 6).Value
                ' calculate the difference between the opening price and closing price
                ws.Cells(SummaryRow, 9).Value = closeprice - openprice
                'apply conditional formatting to cells(SummaryRow,9)
                 If ws.Cells(SummaryRow, 9) < 0 Then
                 ws.Cells(SummaryRow, 9).Interior.ColorIndex = 3
                 Else
                 ws.Cells(SummaryRow, 9).Interior.ColorIndex = 4
                    End If
            
                ' write an if statement to handle anytime openprice is 0
                If openprice <> 0 Then
                    Percentchange = (closeprice - openprice) / openprice
                    ws.Cells(SummaryRow, 10).Value = Percentchange
                Else
                ' insert NA if openprice is 0 and Percentchange cannot be calculated
                    ws.Cells(SummaryRow, 10).Value = "NA"
                End If
            
                'change the value of the opening price for the next ticker
                openprice = ws.Cells(i + 1, 3).Value
            
                TotalTicker = TotalTicker + ws.Cells(i, 7).Value
                ws.Cells(SummaryRow, 11).Value = TotalTicker
                SummaryRow = SummaryRow + 1
                TotalTicker = 0
                
            Else
            TotalTicker = TotalTicker + ws.Cells(i, 7).Value
            
            End If
            'END IF STATEMENT
        
        Next i
    Next ws

End Sub



