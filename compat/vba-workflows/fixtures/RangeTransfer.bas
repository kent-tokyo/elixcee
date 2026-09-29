Attribute VB_Name = "RangeTransfer"
Option Explicit

Public Sub TransferValue2Matrix()
    Range("D1").Formula = "=1+2"
    Cells(1, 5).Value2 = 8
    Cells(2, 4).Value2 = 9
    Range("E2").ClearContents

    Range("A1:B2").Value2 = Range("D1:E2").Value2
End Sub
