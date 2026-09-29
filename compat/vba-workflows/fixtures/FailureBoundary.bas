Attribute VB_Name = "FailureBoundary"
Option Explicit

Public Sub FailsAfterWrite()
    Cells(1, 1).Value = 99
    Err.Raise 5
End Sub
