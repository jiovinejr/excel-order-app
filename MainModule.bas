Attribute VB_Name = "Module1"
Sub AddToOnDeck()

Dim checkRng As Range, shipCheck As Range, onDeckRng As Range
Dim shipName As String, lastInCheck As Integer, lastOnDeck As Integer, target As Integer
Dim r As Variant
Dim qty As Double, meas As String, item As String

shipName = Worksheets("Check").Range("B1").Value
lastInCheck = Worksheets("Check").Range("A" & Rows.Count).End(xlUp).Row
Set checkRng = Worksheets("Check").Range("A4", "C" & lastInCheck)
lastOnDeck = Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row
target = lastOnDeck + 1
Set onDeckRng = Worksheets("On Deck").Range("A2:A" & lastOnDeck)

Set shipCheck = Worksheets("On Deck").Range("A2", "A" & lastOnDeck)

With onDeckRng
    For r = 1 To .Rows.Count
        If .Cells(r, 1) = shipName Then
            .Rows(r).EntireRow.Delete
            r = r - 1
        End If
    Next r
End With

lastOnDeck = Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row
target = lastOnDeck + 1

If Application.CountIf(shipCheck, shipName) = 0 Then
For Each r In checkRng.Rows
    qty = r.Cells(, 1)
    meas = r.Cells(, 2)
    item = r.Cells(, 3)
    
    Worksheets("On Deck").Range("A" & target) = shipName
    Worksheets("On Deck").Range("B" & target) = qty
    Worksheets("On Deck").Range("C" & target) = meas
    Worksheets("On Deck").Range("D" & target) = item
    
    target = target + 1
    
    
Next
End If

FilterDeck

End Sub

Sub FilterDeck()

Dim ws As Worksheet
Dim shipsRng As Range, rngDestination As Range

Set ws = ThisWorkbook.Sheets("On Deck")

Set shipRng = Worksheets("On Deck").Range("F1:F" & Worksheets("On Deck").Range("F" & Rows.Count).End(xlUp).Row)
Set rngDestination = Worksheets("On Deck").Range("F1:F1")

shipRng.Clear
Worksheets("On Deck").Range("A1:A" & Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=rngDestination, Unique:=True

shipRng.Sort Key1:=shipRng.Cells(1), Order1:=xlAscending, Header:=xlYes
ws.AutoFilterMode = False

End Sub


Sub LabelsForFullOrder()
    last = Worksheets("Label") _
        .Range("C" & Rows.Count).End(xlUp).Row

        
    shipName = Worksheets("Label").Range("E1").Text
    
    PrintBoxLabels 1, last
    
    PrintRollLabel (shipName)
End Sub

Sub SelectedLabels()

    Selected = Selection.Areas(1).Rows.Count
    r = CInt(Selection.Areas(1).Cells.Row)
    last = (r + Selected) - 1
    
    PrintBoxLabels r, last
    
End Sub

Sub PrintOrderAndCheck()
    Dim orderRng As Variant, checkRng As Variant
    Dim shipName As String, lastInOrder As Integer
    
    
    shipName = Worksheets("Check").Range("B1")
    
    lastInOrder = Worksheets("Order").Range("A" & Rows.Count).End(xlUp).Row
    
    
    
    Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"
    
    Set orderRng = Worksheets("Order").Range("A1", "E" & lastInOrder)
    Set checkRng = Worksheets("Check").Range("A1", "D" & lastInOrder)
    
    checkRng.PrintOut
    orderRng.PrintOut
End Sub

Sub GetActPrint()

    Dim sheetName As String
    sheetName = ActiveSheet.Name
    MsgBox sheetName
    

End Sub

Sub MakeCheckSheet()
    Dim orderRange As Range
    Dim orderQty As Double, orderMeasurement As String, orderProduct As String
    Dim msrmntLookupRange As Range, prdctLookupRange As Range
    Dim lastInOrder As Integer, orderRow As Variant
    Dim ship As Variant
    

    Set msrmntLookupRange = Worksheets("Master List").Range("F:G")
    Set prdctLookupRange = Worksheets("Master List").Range("B:C")
    

    ship = Worksheets("Order").Range("C1").Value
    

    lastInOrder = Worksheets("Order").Cells(Rows.Count, "C").End(xlUp).Row
    
 
    Set orderRange = Worksheets("Order").Range("A4:C" & lastInOrder)
    
    Worksheets("Check").Range("A4:C150").ClearContents
    Worksheets("Check").Range("A4:C150").Interior.Color = vbWhite
    

    Worksheets("Check").Range("B1").Value = ship
    

    Worksheets("Check").Activate
    

    For Each orderRow In orderRange.Rows
        orderQty = orderRow.Cells(1, 1).Value
        orderMeasurement = orderRow.Cells(1, 2).Value
        orderProduct = orderRow.Cells(1, 3).Value
        

        orderRow.Cells(1, 1).Copy
        Worksheets("Check").Cells(orderRow.Row, 1).PasteSpecial Paste:=xlPasteValues
        
        On Error Resume Next
        Worksheets("Check").Cells(orderRow.Row, 2) = Application.WorksheetFunction.VLookup(orderMeasurement, msrmntLookupRange, 2, False)
        If Err.Number <> 0 Then
            Worksheets("Check").Cells(orderRow.Row, 2) = orderMeasurement
            Worksheets("Check").Cells(orderRow.Row, 2).Interior.Color = vbYellow
            Err.Clear
        End If
        
        Worksheets("Check").Cells(orderRow.Row, 3) = Application.WorksheetFunction.VLookup(orderProduct, prdctLookupRange, 2, False)
        If Err.Number <> 0 Then
            Worksheets("Check").Cells(orderRow.Row, 3) = orderProduct
            Worksheets("Check").Cells(orderRow.Row, 3).Interior.Color = vbYellow
            Err.Clear
        End If
        On Error GoTo 0
    Next orderRow
    
    ' Sort the Check sheet by the third column
    Worksheets("Check").Sort.SortFields.Clear
    Worksheets("Check").Sort.SortFields.Add Key:=Worksheets("Check").Range("C4:C" & lastInOrder), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Worksheets("Check").Sort
        .SetRange Worksheets("Check").Range("A4:C" & lastInOrder)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


Sub Breakdown()

Dim rng As Range, arr() As Variant, targetCell As Range, splitSize As Double
    
    Set targetCell = Worksheets("Label").Range("A1")
    splitSize = 1
    
    With Worksheets("Label")
        If .Range("A1").Value <> "" Then
            .Range("A1:C" & .Range("C" & .Rows.Count).End(xlUp).Row).Clear
        End If
    End With
    
    With Worksheets("Check")
        Set rng = .Range("A4:C" & .Range("C" & .Rows.Count).End(xlUp).Row)
    End With
    arr = rng
    
    Worksheets("Label").Activate
    
    Dim quantity As Double, packaging As String, item As String, rowCounter As Long, caseWeight As Double, i As Integer
    Worksheets("Label").Range("E1").Value = Worksheets("Check").Range("B1")
    
For i = 1 To UBound(arr)
    quantity = arr(i, 1)
    packaging = arr(i, 2)
    item = arr(i, 3)
On Error Resume Next
    caseWeight = Application.WorksheetFunction.XLookup(item, Worksheets("Master List").Range("C:C"), Worksheets("Master List").Range("E:E"))
If Err.Number <> 0 Then
    caseWeight = quantity
    msgString = item & " is not in Master List. Add to Master List and re-process from begining."
    MsgBox msgString
End If
    
    If packaging = "Bag" And item Like "*Radish*" Then
        ProcessBagRadish quantity, packaging, item, targetCell, rowCounter
    ElseIf item Like "*Watermelon*" Then
        ProcessWatermelon quantity, packaging, item, targetCell, rowCounter, caseWeight
    ElseIf packaging = "Bunch" Then
        ProcessBunch quantity, packaging, item, targetCell, rowCounter
    ElseIf packaging <> "Pound" Then
        ProcessNonPound quantity, packaging, item, targetCell, rowCounter, splitSize
    Else
        ProcessPound quantity, packaging, item, targetCell, rowCounter, caseWeight
    End If
Next i

AddToOnDeck
FilterDeck
RefreshOnDeckPivot

End Sub

Sub ProcessBagRadish(quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    While quantity > 30
        WriteLabel 30, packaging, item, targetCell, rowCounter
        quantity = quantity - 30
    Wend
    WriteLabel quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessWatermelon(quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, caseWeight As Double)
    While quantity > caseWeight
        WriteLabel "", packaging, item, targetCell, rowCounter
        quantity = quantity - caseWeight
    Wend
    WriteLabel "", packaging, item, targetCell, rowCounter
End Sub

Sub ProcessBunch(quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    While quantity > 48
        WriteLabel 48, packaging, item, targetCell, rowCounter
        quantity = quantity - 48
    Wend
    WriteLabel quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessNonPound(quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, splitSize As Double)
    While quantity > splitSize
        WriteLabel splitSize, packaging, item, targetCell, rowCounter
        quantity = quantity - splitSize
    Wend
    WriteLabel quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessPound(quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, caseWeight As Double)
    While quantity > caseWeight
        WriteLabel caseWeight, packaging, item, targetCell, rowCounter
        quantity = quantity - caseWeight
    Wend
    WriteLabel quantity, packaging, item, targetCell, rowCounter
End Sub

Sub WriteLabel(quantity As Variant, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    ' Write label information to the target cell and increment the row counter
    targetCell.Offset(rowCounter, 0).Value = quantity
    targetCell.Offset(rowCounter, 1).Value = packaging
    targetCell.Offset(rowCounter, 2).Value = item
    rowCounter = rowCounter + 1
End Sub

Sub PasteSpeacial()

On Error Resume Next
    Worksheets("Order").Range("A1:C200").ClearContents
    Worksheets("Order").Range("A1").Select
    Worksheets("Order").PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    
If Err.Number <> 0 Then
    MsgBox ("Nothing copied. Copy something, will ya!")
End If

End Sub
