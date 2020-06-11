Attribute VB_Name = "Module1"
Sub stocks()

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Svolume As Double
Svolume = 0

Dim Sticker As Integer
Sticker = 2

Dim ClosePrice As Double
Dim OpenPrice As Double
Dim Diff As Double

For i = 2 To lastrow

    If OpenPrice = 0 Then
    OpenPrice = ws.Cells(i, 3).Value
    
    ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    Svolume = Svolume + ws.Cells(i, 7).Value
    
    Else
    Svolume = Svolume + ws.Cells(i, 7).Value
    
    ws.Cells(Sticker, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(Sticker, 12).Value = Svolume
    ClosePrice = ws.Cells(i, 6).Value
    Diff = (ClosePrice - OpenPrice)
    ws.Cells(Sticker, 10).Value = Diff
    ws.Cells(Sticker, 11).Value = (Diff / OpenPrice)
    ws.Cells(Sticker, 11).NumberFormat = "0.00%"
    
    Svolume = 0
    Sticker = Sticker + 1
    OpenPrice = 0
    ClosePrice = 0
    Diff = 0
    
    End If
Next i


For i = 2 To lastrow
          
    If ws.Cells(i, 11).Value < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
    
    End If
Next i

ws.Columns("I:L").AutoFit

Next ws
    
End Sub
