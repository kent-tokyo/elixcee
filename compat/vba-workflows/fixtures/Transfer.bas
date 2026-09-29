Attribute VB_Name = "Transfer"
Option Explicit

Public Sub TransferAndRecalculate()
    Worksheets("Sheet1").Cells(2, 2).Value = 30
    Worksheets("Sheet1").Cells(3, 2).Value = 60
End Sub
