Attribute VB_Name = "ObjectAliases"
Option Explicit

Private items As Collection
Private values As Scripting.Dictionary

Public Sub RunObjectAliasCase()
    Set items = New Collection
    items.Add 4
    Set values = New Scripting.Dictionary
    values.Add "first", 1

    Call AddThroughAliases

    Cells(2, 2).Value = items.Count
    Cells(3, 2).Value = values.Count
    Cells(4, 2).Value = items.Item(2) + values.Item("second")
End Sub

Private Sub AddThroughAliases()
    Dim itemAlias As Collection
    Dim valueAlias As Scripting.Dictionary
    Set itemAlias = items
    Set valueAlias = values
    itemAlias.Add 8
    valueAlias("second") = 2
End Sub
