Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Public Sub PrintFile(ByVal strPathAndFilename As String)

   Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"
   Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)

End Sub

Sub GoToOrder()

Worksheets("Order").Activate

End Sub
Sub MultiSkid()

Dim labelPath As String, skids As Variant, i As Integer, t As String, numOfSkids As Integer
Dim ObjDoc As bpac.Document
Set ObjDoc = CreateObject("bpac.Document")

labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeMulti.lbx"
skids = InputBox("How many skids?", "MultiSkid", "2")

If skids <> "" Then
numOfSkids = CInt(skids)

ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoCutAtEnd
        
        For i = 1 To numOfSkids
        
            t = i & " of " & numOfSkids
            
            ObjDoc.GetObject("Multi").Text = t
            ObjDoc.PrintOut 2, bpoDefault
        
        Next i
    
    ObjDoc.EndPrint
    ObjDoc.Close
End If


End Sub

Sub AddToOnDeck()
Dim checkRng As Range, shipCheck As Range, onDeckRng As Range
Dim shipName As String, lastInCheck As Integer, lastOnDeck As Integer, target As Integer
Dim r As Variant
Dim qty As Double, meas As String, item As String

shipName = Worksheets("Check").Range("B1").value
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
Dim shipRng As Range, rngDestination As Range

Set ws = ThisWorkbook.Sheets("On Deck")

Set shipRng = Worksheets("On Deck").Range("F1:F" & Worksheets("On Deck").Range("F" & Rows.Count).End(xlUp).Row)
Set rngDestination = Worksheets("On Deck").Range("F1:F1")

shipRng.Clear
Worksheets("On Deck").Range("A1:A" & Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=rngDestination, Unique:=True

shipRng.Sort Key1:=shipRng.Cells(1), Order1:=xlAscending, Header:=xlYes
ws.AutoFilterMode = False

Worksheets("On Deck").Range("A1:D1").AutoFilter

End Sub
Sub PrintBoxLabels(begin As Variant, last As Variant)
Dim sheetName As String, labelPath As String, shipName As String, i As Integer

Dim ObjDoc As bpac.Document, kg As Double
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeCaseLabels2.lbx"
    shipName = Worksheets("Label").Range("E1").Text
    
  
    sheetName = ActiveSheet.Name
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoCutAtEnd

Dim qty As String, meas As String, item As String
            For i = begin To last
                kg = Format(Range("A" & i) / 2.2, "0.00")
                qty = Range("A" & i).Text
                meas = Range("B" & i).Text
                item = Range("C" & i).Text
                
                If sheetName <> "Label" Then
                    shipName = Range("D" & i).Text
                End If
                
                ObjDoc.GetObject("DelShip").Text = "Delaware Ship Supply Co."
                
                ObjDoc.GetObject("Ship").Text = shipName
                
                ObjDoc.GetObject("Qty").Text = qty
                
                ObjDoc.GetObject("Measure").Text = meas
                
                ObjDoc.GetObject("Item").Text = item
                
                If kg <> 0 Then
                    ObjDoc.GetObject("Kilo").Text = "(" & kg & " Kilo)"
                Else
                    ObjDoc.GetObject("Kilo").Text = ""
                End If
                
                ObjDoc.PrintOut 1, bpoDefault
            Next i
            
            
        ObjDoc.EndPrint
    ObjDoc.Close
    
End Sub

Sub LabelsForFullOrder()
Dim l As Integer, shipName As String
    l = Worksheets("Label") _
        .Range("C" & Rows.Count).End(xlUp).Row
        
    shipName = Worksheets("Label").Range("E1").Text
    
    PrintBoxLabels 1, l
    
    PrintRollLabel
End Sub

Sub SelectedLabels()

    Dim Selected As Integer, r As Integer, last As Integer

    Selected = Selection.Areas(1).Rows.Count
    r = CInt(Selection.Areas(1).Cells.Row)
    last = (r + Selected) - 1
    
    PrintBoxLabels r, last
    
End Sub

Sub PrintOrderAndCheck()
    Dim orderRng As Variant, checkRng As Variant, mainFolder As String
    Dim shipName As String, lastInOrder As Integer, ship As String, filePath As String
    
    ship = Worksheets("Label").Range("E1")
    
    mainFolder = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\"
    filePath = mainFolder & ship & "\" & ship
    
    shipName = Worksheets("Check").Range("B1")
    
    lastInOrder = Worksheets("Order").Range("A" & Rows.Count).End(xlUp).Row
    
    Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"
    
    Set orderRng = Worksheets("Order").Range("A1", "E" & lastInOrder)
    Set checkRng = Worksheets("Check").Range("A1", "D" & lastInOrder)
    
    If ship = shipName Then
        checkRng.PrintOut
        orderRng.PrintOut
    Else
        
        PrintFile filePath & "-check.pdf"
        Application.Wait (Now + TimeValue("0:00:04"))
        PrintFile filePath & "-order.pdf"
        
        
    End If
    
    '"C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\MV GRAND PIONEER-328995\MV GRAND PIONEER-328995-order.pdf"
    
End Sub

Sub MakePDFs()
    Dim MyObj As Object, MySource As Object, file As Variant, shipName As String, lastInOrder As Integer
    Dim orderRng As Variant, checkRng As Variant, mainFolder As String, newFolder As String, orderFileName As String, checkFileName As String
    
    'Application.ActivePrinter = "ET-5880 Series(Network) on Ne05:"
    shipName = Worksheets("Check").Range("B1")
    
    lastInOrder = Worksheets("Order").Range("A" & Rows.Count).End(xlUp).Row
    
    Set orderRng = Worksheets("Order").Range("A1", "E" & lastInOrder)
    Set checkRng = Worksheets("Check").Range("A1", "D" & lastInOrder)
    
    mainFolder = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\"
    
    newFolder = mainFolder & shipName & "\"
     
    On Error Resume Next
    MkDir newFolder
    
    orderFileName = shipName & "-order"
    checkFileName = shipName & "-check"
    
    orderRng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newFolder & orderFileName, IgnorePrintAreas:=False
    checkRng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newFolder & checkFileName, IgnorePrintAreas:=False
End Sub



Sub GetActPrint()

    Dim sheetName As String
    sheetName = ActiveSheet.Name
    MsgBox sheetName
    

End Sub


Sub MakeCheckSheet()
    Dim orderRange As Range
    Dim orderQty As Double, OrderMeasurement As String, orderProduct As String
    Dim msrmntLookupRange As Range, prdctLookupRange As Range
    Dim lastInOrder As Integer, orderRow As Variant
    Dim ship As Variant
    
    ' Define the lookup ranges
    Set msrmntLookupRange = Worksheets("Master List").Range("F:G")
    Set prdctLookupRange = Worksheets("Master List").Range("B:C")
    
    ' Get the shipping information
    ship = Worksheets("Order").Range("C1").value
    
    ' Find the last row in the order sheet
    lastInOrder = Worksheets("Order").Cells(Rows.Count, "C").End(xlUp).Row
    
    ' Set the order range
    Set orderRange = Worksheets("Order").Range("A4:C" & lastInOrder)
    
    ' Clear the Check sheet
    Worksheets("Check").Range("A4:C150").ClearContents
    Worksheets("Check").Range("A4:C150").Interior.Color = vbWhite
    
    ' Set the shipping information
    Worksheets("Check").Range("B1").value = ship
    
    'Switch to Check sheet
    Worksheets("Check").Activate
    
    ' Process the orders
    For Each orderRow In orderRange.Rows
        orderQty = orderRow.Cells(1, 1).value
        OrderMeasurement = orderRow.Cells(1, 2).value
        orderProduct = orderRow.Cells(1, 3).value
        
        ' Write order information to Check sheet
        orderRow.Cells(1, 1).Copy
        Worksheets("Check").Cells(orderRow.Row, 1).PasteSpecial Paste:=xlPasteValues
        ' Lookup measurement and product information
        On Error Resume Next
        Worksheets("Check").Cells(orderRow.Row, 2) = Application.WorksheetFunction.VLookup(OrderMeasurement, msrmntLookupRange, 2, False)
        If Err.Number <> 0 Then
            Worksheets("Check").Cells(orderRow.Row, 2) = OrderMeasurement
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
    Worksheets("Check").Sort.SortFields.Add Key:=Worksheets("Check").Range("C4:C" & lastInOrder), SortOn:=xlSortOnValues, ORDER:=xlAscending, DataOption:=xlSortNormal
    With Worksheets("Check").Sort
        .SetRange Worksheets("Check").Range("A4:C" & lastInOrder)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Print the Order sheet
   ' Worksheets("Order").Range("A1:E" & lastInOrder).PrintPreview
End Sub




Sub Breakdown()

Dim rng As Range, arr() As Variant, targetCell As Range, splitSize As Double
    
    'targetCell is where the result will start being written
    Set targetCell = Worksheets("Label").Range("A1")
    splitSize = 1
    
    'clear the previous labels, if there are any
    With Worksheets("Label")
        If .Range("A1").value <> "" Then
            .Range("A1:C" & .Range("C" & .Rows.Count).End(xlUp).Row).Clear
        End If
    End With
    
    'find the last row on the "order" sheet, assign the range with the order to the arr array
    With Worksheets("Check")
        Set rng = .Range("A4:C" & .Range("C" & .Rows.Count).End(xlUp).Row)
    End With
    arr = rng
    
    'Switch to Label sheet
    Worksheets("Label").Activate
    
    Dim Quantity As Double, packaging As String, item As String, rowCounter As Long, caseWeight As Double, i As Integer
    Worksheets("Label").Range("E1").value = Worksheets("Check").Range("B1")
    
' Loop through all rows of the order
For i = 1 To UBound(arr)
    Quantity = arr(i, 1)
    packaging = arr(i, 2)
    item = arr(i, 3)
On Error Resume Next
    caseWeight = Application.WorksheetFunction.XLookup(item, Worksheets("Master List").Range("C:C"), Worksheets("Master List").Range("E:E"))
If Err.Number <> 0 Then
    caseWeight = Quantity
    msgString = item & " is not in Master List. Add to Master List and re-process from begining."
    MsgBox msgString
End If
    
    If packaging = "Bag" And item Like "*Radish*" Then
        ProcessBagRadish Quantity, packaging, item, targetCell, rowCounter
    ElseIf item Like "*Watermelon*" Then
        ProcessWatermelon Quantity, packaging, item, targetCell, rowCounter, caseWeight
    ElseIf packaging = "Bunch" Or packaging = "Each" Then
        ProcessBunch Quantity, packaging, item, targetCell, rowCounter
    ElseIf packaging <> "Pound" Then
        ProcessNonPound Quantity, packaging, item, targetCell, rowCounter, splitSize
    Else
        ProcessPound Quantity, packaging, item, targetCell, rowCounter, caseWeight
    End If
Next i

AddToOnDeck
FilterDeck
RefreshOnDeckPivot

End Sub

Sub ProcessBagRadish(Quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    While Quantity > 30
        WriteLabel 30, packaging, item, targetCell, rowCounter
        Quantity = Quantity - 30
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessWatermelon(Quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, caseWeight As Double)
    While Quantity > caseWeight
        WriteLabel "", packaging, item, targetCell, rowCounter
        Quantity = Quantity - caseWeight
    Wend
    WriteLabel "", packaging, item, targetCell, rowCounter
End Sub

Sub ProcessBunch(Quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    While Quantity > 48
        WriteLabel 48, packaging, item, targetCell, rowCounter
        Quantity = Quantity - 48
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessNonPound(Quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, splitSize As Double)
    While Quantity > splitSize
        WriteLabel splitSize, packaging, item, targetCell, rowCounter
        Quantity = Quantity - splitSize
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub ProcessPound(Quantity As Double, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long, caseWeight As Double)
    While Quantity > caseWeight
        WriteLabel caseWeight, packaging, item, targetCell, rowCounter
        Quantity = Quantity - caseWeight
    Wend
    WriteLabel Quantity, packaging, item, targetCell, rowCounter
End Sub

Sub WriteLabel(Quantity As Variant, packaging As String, item As String, targetCell As Range, ByRef rowCounter As Long)
    ' Write label information to the target cell and increment the row counter
    targetCell.Offset(rowCounter, 0).value = Quantity
    targetCell.Offset(rowCounter, 1).value = packaging
    targetCell.Offset(rowCounter, 2).value = item
    rowCounter = rowCounter + 1
End Sub

Sub PrintSkidLabel()
    
    Dim labelPath As String, shipName As String

    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeSkidLabel.lbx"
    shipName = Worksheets("Label").Range("E1").Text
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("ShipName").Text = shipName
            ObjDoc.PrintOut 1, bpoDefault
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close

    
End Sub



Sub PrintRollLabel()

    Dim ObjDoc As bpac.Document, labelPath As String, shipSend As String
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\Protected Folder - DO NOT DELETE\ZeeRollLabel.lbx"
    shipSend = Worksheets("Label").Range("E1").Text

    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("RollLabel").Text = shipSend
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close
 
 

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
