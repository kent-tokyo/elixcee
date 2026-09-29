Attribute VB_Name = "TableQuality"
Option Explicit

Public Sub SortFilterAndFind()
    Cells(1, 1).Value = "Name"
    Cells(1, 2).Value = "Age"
    Cells(2, 1).Value = "Alice"
    Cells(2, 2).Value = 40
    Cells(3, 1).Value = "Bob"
    Cells(3, 2).Value = 10
    Cells(4, 1).Value = "Alice"
    Cells(4, 2).Value = 25

    Range("A1:B4").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
    Range("A1:B4").AutoFilter Field:=2, Criteria1:="25"

    Cells(1, 4).Value = Cells.Find(What:="Alice", MatchCase:=True).Row
    Cells(2, 4).Value = Cells(2, 1).Value
    Cells(3, 4).Value = Cells(3, 1).Value
End Sub
