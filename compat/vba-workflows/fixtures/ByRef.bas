Attribute VB_Name = "ByRef"
Option Explicit

Public Sub AddInto(ByRef value As Integer, ByVal amount As Integer)
    value = value + amount
End Sub

Public Sub RunByRefCase()
    Dim total As Integer
    total = 5
    Call AddInto(total, 7)
    Cells(2, 2).Value = total
    Cells(3, 2).Value = 7
End Sub
