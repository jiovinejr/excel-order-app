Attribute VB_Name = "Module2"

Sub DisplayUserForm()
    
    UserForm1.Show
    
End Sub

Sub GetTodaysList()

Dim lastTodayListItem As Integer, lastShipForToday As Integer, lastOnDeck As Long, target As Integer, lastWeekItem As Integer, weekTarget As Integer
Dim shipsForToday As Range, onDeckRng As Range
Dim r As Variant
Dim shipName As String, qty As Double, meas As String, item As String

lastTodayListItem = Worksheets("Daily").Range("A" & Rows.Count).End(xlUp).Row
lastWeekItem = Worksheets("Week").Range("A" & Rows.Count).End(xlUp).Row
lastShipForToday = Worksheets("Daily").Range("F" & Rows.Count).End(xlUp).Row
lastOnDeck = Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row

Set shipsForToday = Worksheets("Daily").Range("F2:F" & lastShipForToday)
Set onDeckRng = Worksheets("On Deck").Range("A2:A" & lastOnDeck)

target = lastTodayListItem + 1
weekTarget = lastWeekItem + 1

        With onDeckRng
            For r = 1 To .Rows.Count
                shipName = .Cells(r, 1).Value
                qty = .Cells(r, 2).Value
                meas = .Cells(r, 3).Value
                item = .Cells(r, 4).Value
                
                If Application.CountIf(shipsForToday, shipName) > 0 Then
                    Worksheets("Daily").Range("A" & target) = qty
                    Worksheets("Week").Range("A" & weekTarget) = qty
                    Worksheets("Daily").Range("B" & target) = meas
                    Worksheets("Week").Range("B" & weekTarget) = meas
                    Worksheets("Daily").Range("C" & target) = item
                    Worksheets("Week").Range("C" & weekTarget) = item
                    Worksheets("Daily").Range("D" & target) = shipName
                    Worksheets("Week").Range("D" & weekTarget) = shipName
                    .Rows(r).EntireRow.Delete
                    r = r - 1
                    
                    target = target + 1
                    weekTarget = weekTarget + 1
                    
                End If
            Next r
        End With

FilterDeck
RefreshPivot
RefreshOnDeckPivot

End Sub

Sub RefreshPivot()
If Worksheets("Daily").Range("C2") > "" Then
    Worksheets("Needs").PivotTables("Day").RefreshTable
End If

End Sub

Sub RefreshOnDeckPivot()
Worksheets("Items on Deck").PivotTables("ItemsOnDeck").RefreshTable
End Sub

Sub PrintPivot()
Dim printRng As Range
    last = Worksheets("Needs").Range("A" & Rows.Count).End(xlUp).Row
    
    Set printRng = Worksheets("Needs").Range("A1:G" & last)
    
    Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"

    printRng.PrintOut
    
End Sub

Sub ClearDaily()
Dim lastToday As Integer

lastToday = Worksheets("Daily").Range("A" & Rows.Count).End(xlUp).Row


Worksheets("Daily").Range("A2:F3000").ClearContents


End Sub
