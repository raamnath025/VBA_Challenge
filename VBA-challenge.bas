Attribute VB_Name = "Module1"
Option Explicit
Sub StockData()
    Const FIRST_DATA_ROW As Integer = 2
    Const OPEN_COL As Integer = 3
    Const CLOSE_COL As Integer = 6
    Const VOL_COL As Integer = 7
    
        Dim WS As Worksheet
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
          Volume = 0
        Dim Column As Integer
          Column = 1
        Dim inputRow As Long
        Dim outputRow As Long
        Dim CP As Double    'CP = Close Price
        Dim PO As Double    'PO = Open Price
        Dim LastRow As Double
          For Each WS In ActiveWorkbook.Worksheets
                  WS.Activate
           LastRow = Cells(Rows.Count, 1).End(xlUp).Row
           outputRow = FIRST_DATA_ROW
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
          
        PO = Cells(FIRST_DATA_ROW, OPEN_COL).Value
        For inputRow = FIRST_DATA_ROW To LastRow
            If Cells(inputRow + 1, Column).Value <> Cells(inputRow, Column).Value Then
                Ticker_Name = Cells(inputRow, Column).Value
                Volume = Volume + Cells(inputRow, VOL_COL).Value
                
                CP = Cells(inputRow, CLOSE_COL).Value
                Yearly_Change = CP - PO
                If (PO = 0) Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / PO
                End If
                
                Cells(outputRow, Column + 9).Value = Yearly_Change
                Cells(outputRow, Column + 8).Value = Ticker_Name
                Cells(outputRow, Column + 10).Value = Percent_Change
                Cells(outputRow, Column + 10).NumberFormat = "0.00%"
                Cells(outputRow, Column + 11).Value = Volume
                'Prepare for next stock
                outputRow = outputRow + 1
                PO = Cells(inputRow + 1, OPEN_COL)
                Volume = 0
            Else
                Volume = Volume + Cells(inputRow, VOL_COL).Value
            End If
        Next inputRow
        
        Dim RGLastRow As Double
        RGLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        Dim j As Integer
        For j = 2 To RGLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        'Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        'Look through each rows to find the greatest value and its associate ticker
        Dim GR As Integer
        For GR = 2 To RGLastRow
            If Cells(GR, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & RGLastRow)) Then
                Cells(2, Column + 15).Value = Cells(GR, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(GR, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(GR, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & RGLastRow)) Then
                Cells(3, Column + 15).Value = Cells(GR, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(GR, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(GR, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & RGLastRow)) Then
                Cells(4, Column + 15).Value = Cells(GR, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(GR, Column + 11).Value
            End If
        Next GR
    Next WS

End Sub
