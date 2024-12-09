VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickerloop()

    Dim ws As Worksheet
    Dim last_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim quaterly_change As Double
    Dim ticker As String
    Dim percent_change As Double
    Dim volume As Double
    Dim row As Long
    Dim column As Integer

    For Each ws In ActiveWorkbook.Worksheets
    
        ' Find the last row in column A
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

        ' Add headers
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With

        ' Initialize variables
        volume = 0
        row = 2
        column = 1
        open_price = ws.Cells(2, column + 2).Value
        
        ' Loop through rows
        For i = 2 To last_row
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ' Set ticker name
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, 9).Value = ticker
                
                ' Calculate close price and quarterly change
                close_price = ws.Cells(i, column + 5).Value
                quaterly_change = close_price - open_price
                ws.Cells(row, 10).Value = quaterly_change
                
                ' Calculate percent change
                If open_price <> 0 Then
                    percent_change = quaterly_change / open_price
                Else
                    percent_change = 0
                End If
                ws.Cells(row, 11).Value = percent_change
                ws.Cells(row, 11).NumberFormat = "0.00%"
                
                ' Total stock volume
                volume = volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, 12).Value = volume
                
                ' Move to next row
                row = row + 1
                
                ' Reset open price and volume
                If i + 1 <= last_row Then
                    open_price = ws.Cells(i + 1, column + 2).Value
                End If
                volume = 0
            Else
                volume = volume + ws.Cells(i, column + 6).Value
            End If
        Next i

        ' Apply conditional formatting for Quarterly Change
        Dim quaterly_change_last_row As Long
        quaterly_change_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        For j = 2 To quaterly_change_last_row
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10 ' Green for positive
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red for negative
            End If
        Next j

        ' Find max and min for Percent Change and Volume
        Dim max_percent_row As Long, min_percent_row As Long, max_volume_row As Long
        max_percent_row = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)), ws.Range("K2:K" & quaterly_change_last_row), 0) + 1
        min_percent_row = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row)), ws.Range("K2:K" & quaterly_change_last_row), 0) + 1
        max_volume_row = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row)), ws.Range("L2:L" & quaterly_change_last_row), 0) + 1

        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 17).Value = ws.Cells(max_percent_row, 9).Value
        ws.Cells(2, 18).Value = ws.Cells(max_percent_row, 11).Value
        ws.Cells(2, 18).NumberFormat = "0.00%"

        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 17).Value = ws.Cells(min_percent_row, 9).Value
        ws.Cells(3, 18).Value = ws.Cells(min_percent_row, 11).Value
        ws.Cells(3, 18).NumberFormat = "0.00%"

        ws.Cells(4, 16).Value = "Highest Volume"
        ws.Cells(4, 17).Value = ws.Cells(max_volume_row, 9).Value
        ws.Cells(4, 18).Value = ws.Cells(max_volume_row, 12).Value

        ' Format columns
        With ws.Range("I:Q")
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
    Next ws

End Sub

