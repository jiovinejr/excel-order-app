Attribute VB_Name = "Reports"
'Create Reports

'Create check sheet from an order array
Sub CreateCheckSheet(arr() As OrderRecord)

'Initialize
Dim sortedArr() As OrderRecord, checkSheet As Worksheet
Dim checkRange As Range, checkShipName As Range

'Sort the incoming array
sortedArr = SortOrderRecord(arr)

'Set excel poits to variables
Set checkSheet = Worksheets("CheckPrint")
Set checkShipName = checkSheet.Range("B1")
Set checkRange = checkSheet.Range("A4:C" & UBound(sortedArr))

'Clear anything and set headers
checkSheet.Cells.ClearContents
checkSheet.Range("A1").value = "Name:"
checkSheet.Range("A2").value = "Date:"

'Set ship name
checkShipName.value = sortedArr(1).ship

'Incrementor
Dim i As Integer
i = 1

'Loop through array and write data to cells
For Each ordRec In sortedArr
    checkRange.Cells(i, 1) = ordRec.Quantity
    checkRange.Cells(i, 2) = ordRec.CleanMeasurement
    checkRange.Cells(i, 3) = ordRec.CleanItem
    'Next row
    i = i + 1
Next ordRec

'Hide sheet
checkSheet.Visible = xlSheetHidden

End Sub

'Create an order sheet from an order array
Sub CreateOrderSheet(arr() As OrderRecord)

'Initialize
Dim orderSheet As Worksheet
Dim orderRange As Range, orderShipName As Range

'Set excel points to variables
Set orderSheet = Worksheets("OrderPrint")
Set orderShipName = orderSheet.Range("C1")
Set orderRange = orderSheet.Range("A4:C" & UBound(arr))

'Clear out sheet
orderSheet.Cells.ClearContents

'Set the ship name
orderShipName.value = arr(1).ship

'Incrementor
Dim i As Integer
i = 1

'Loop through un-sorted array
For Each ordRec In arr
    orderRange.Cells(i, 1) = ordRec.Quantity
    orderRange.Cells(i, 2) = ordRec.OrderMeasurement
    orderRange.Cells(i, 3) = ordRec.OrderItem
    'Next row
    i = i + 1
Next ordRec

'Hide sheet
orderSheet.Visible = xlSheetHidden

End Sub

'Sub to check subs in this Mod
Sub CheckReportTest()
Dim orderArr() As OrderRecord
orderArr = CreateRecordFromPaste
CreateCheckSheet orderArr
CreateOrderSheet orderArr
End Sub
