Sub testing()
        Dim ws As Worksheet
        Dim ticker As String
        Dim vol As Double
        Dim year_open As Double
        Dim year_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim Summary_Table_Row As Integer
        Dim last_row As Long
        Dim Start As Double
        
        Summary_Table_Row = 2
        
        For Each ws In ThisWorkbook.Worksheets
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            Start = 2
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Summary_Table_Row = 2
            
'                vol = 0
'                year_open = 0
'                year_close = 0
        
                 year_open = ws.Cells(2, 3).Value
                
        For i = 2 To last_row
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                'vol = ws.Cells(i, 7).Value
                vol = vol + ws.Cells(i, 7).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                
                year_open = ws.Cells(i + 1, 3).Value
                
                percent_change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                
                If ws.Cells(Start, 3) = 0 Then
                percent_change = 0
                Else
                percent_change = Round((percent_change / ws.Cells(Start, 3)) * 100, 2)
                End If
                                
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                ws.Cells(Summary_Table_Row, 11).Value = percent_change
                'ws.Cells(i, 11).NumberFormat = "0.00%"
                
                If ws.Cells(Summary_Table_Row, 11).Value >= 0 Then
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                End If
                
                ws.Cells(Summary_Table_Row, 11).Value = percent_change & "%"
                ws.Cells(Summary_Table_Row, 12).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1
                
                vol = 0
                yearly_change = 0
                percent_change = 0
                year_close = 0
                year_open = 0
                
                Start = i + 1
                
            Else
                vol = vol + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws



End Sub



