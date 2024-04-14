Attribute VB_Name = "ReadDataFileCOMMENT"
'Read Data File
Public Function MapToOrderRecord(orderRange As Range, ship As String) As Variant

'Initialize variables we will use
Dim ordRec As OrderRecord, arr() As OrderRecord
Dim i As Integer, cnt As Integer

'Initiate a variable to help with array initialization
cnt = orderRange.Rows.Count

'Re-dimension array with proper size
ReDim arr(0 To cnt - 1) As OrderRecord

'Incrementor
i = 0

'Loop through range rows and construct an OrderRecord with data read in
For Each orderRow In orderRange.Rows
    Set ordRec = New OrderRecord
    ordRec.Quantity = orderRow.Cells(, 1).value
    ordRec.OrderMeasurement = orderRow.Cells(, 2).value
    ordRec.OrderItem = orderRow.Cells(, 3).value
    ordRec.ship = ship
    'Add order to array and increment
    Set arr(i) = ordRec
    i = i + 1
    'Debug.Print (ordRec.toString)
Next orderRow

'Return Statement
MapToOrderRecord = arr

End Function

'Bubble Sort order by clean name
Public Function SortOrderRecord(arr() As OrderRecord)

'Initialize
Dim i As Long, j As Long, temp As OrderRecord

'Bubble sort algorithm
'Loop array and if the item that comes first is higher alphabetically than the one that comes after, switch them
For i = LBound(arr) To UBound(arr) - 1
    For j = i + 1 To UBound(arr)
        If arr(i).CleanItem > arr(j).CleanItem Then
            Set temp = arr(i)
            Set arr(i) = arr(j)
            Set arr(j) = temp
        End If
        'Debug.Print (coll(i).CleanItem)
    Next j
Next i

'Return statement
SortOrderRecord = arr
End Function

'Read data from order pasted in
Public Function CreateRecordFromPaste() As Variant

'Initialize variables
Dim orderRange As Range, lastInOrder As Integer
Dim shipFromOrderPaste As String, arr() As OrderRecord

'Get last row with data in the item column "C"
lastInOrder = Worksheets("Order").Cells(Rows.Count, "C").End(xlUp).Row

'Set the range as order starting at the 4th row and going to last row with data
Set orderRange = Worksheets("Order").Range("A4:C" & lastInOrder)

'Define the name of the ship
shipFromOrderPaste = Worksheets("Order").Range("C1").value

'Use mapping function to create an array of OrderRecords
arr = MapToOrderRecord(orderRange, shipFromOrderPaste)

'Return that arr
CreateRecordFromPaste = arr

End Function

Public Function CreateRecordFromDB(shipName As String) As Variant
Dim db As Worksheet, allShipsRange As Range, arr() As OrderRecord
Dim startRowOfOrder As Integer, numOfItems As Integer, orderRange As Range
Dim lastRowOfOrder As Integer
Set db = Worksheets("OrderDatabase")
Set allShipsRange = db.Range("G:G")

numOfItems = Application.WorksheetFunction.XLookup(shipName, Worksheets("ShipDatabase").Range("A:A"), Worksheets("ShipDatabase").Range("B:B"))

startRowOfOrder = allShipsRange.Find(shipName).Row
lastRowOfOrder = startRowOfOrder + (numOfItems - 1)
Set orderRange = db.Range("A" & startRowOfOrder & ":G" & lastRowOfOrder)

arr = MapToOrderRecord(orderRange, shipName)

CreateRecordFromDB = arr
End Function

Sub GetOrderFromDBTest()
Dim ship As String
ship = "MV FLORETGRACHT-329085"
CreateRecordFromDB ship
End Sub

'Quick sub to check the array functions in this Module
Sub PrintOrder()
Dim orderArr() As OrderRecord, sortedArr() As OrderRecord, arr() As OrderRecord
Dim ship As String
ship = "MV FLORETGRACHT-329085"
arr = CreateRecordFromDB(ship)
orderArr = CreateRecordFromPaste
sortedArr = SortOrderRecord(arr)
'For Each rec In orderArr
'    Debug.Print rec.toString
'Next rec
For Each rec In sortedArr
    Debug.Print rec.toString
Next rec
End Sub
