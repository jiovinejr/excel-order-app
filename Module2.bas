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

'DID THIS
        With onDeckRng
            For r = 1 To .Rows.Count
                shipName = .Cells(r, 1).value
                qty = .Cells(r, 2).value
                meas = .Cells(r, 3).value
                item = .Cells(r, 4).value
                
                If Application.CountIf(shipsForToday, shipName) > 0 Then
                    Worksheets("Daily").Range("A" & target) = qty
                    'Worksheets("Week").Range("A" & weekTarget) = qty
                    Worksheets("Daily").Range("B" & target) = meas
                    'Worksheets("Week").Range("B" & weekTarget) = meas
                    Worksheets("Daily").Range("C" & target) = item
                    'Worksheets("Week").Range("C" & weekTarget) = item
                    Worksheets("Daily").Range("D" & target) = shipName
                    'Worksheets("Week").Range("D" & weekTarget) = shipName
                    .Rows(r).EntireRow.Delete
                    r = r - 1
                    
                    target = target + 1
                    weekTarget = weekTarget + 1
                    
                End If
            Next r
        End With
'TO THIS
FilterDeck
RefreshPivot
RefreshOnDeckPivot

End Sub

Sub PutBackOnDeck()
Dim lastTodayListItem As Integer, lastShipForToday As Integer, lastOnDeck As Long, target As Integer, lastWeekItem As Integer, weekTarget As Integer
Dim shipsForToday As Range, onDeckRng As Range, dailyRange As Range
Dim r As Variant
Dim shipName As String, qty As Double, meas As String, item As String

lastTodayListItem = Worksheets("Daily").Range("A" & Rows.Count).End(xlUp).Row
lastWeekItem = Worksheets("Week").Range("A" & Rows.Count).End(xlUp).Row
lastShipForToday = Worksheets("Daily").Range("F" & Rows.Count).End(xlUp).Row
lastOnDeck = Worksheets("On Deck").Range("A" & Rows.Count).End(xlUp).Row

Set shipsForToday = Worksheets("Daily").Range("F2:F" & lastShipForToday)
Set onDeckRng = Worksheets("On Deck").Range("A2:A" & lastOnDeck)
Set dailyRange = Worksheets("Daily").Range("A2:A" & lastTodayListItem)

target = lastOnDeck + 1
weekTarget = lastWeekItem + 1

        With onDeckRng
            For r = 1 To .Rows.Count
                shipName = .Cells(r, 1).value
                qty = .Cells(r, 2).value
                meas = .Cells(r, 3).value
                item = .Cells(r, 4).value
                
                If Application.CountIf(shipsForToday, shipName) > 0 Then
                    Worksheets("Daily").Range("A" & target) = qty
                    'Worksheets("Week").Range("A" & weekTarget) = qty
                    Worksheets("Daily").Range("B" & target) = meas
                    'Worksheets("Week").Range("B" & weekTarget) = meas
                    Worksheets("Daily").Range("C" & target) = item
                    'Worksheets("Week").Range("C" & weekTarget) = item
                    Worksheets("Daily").Range("D" & target) = shipName
                    'Worksheets("Week").Range("D" & weekTarget) = shipName
                    .Rows(r).EntireRow.Delete
                    r = r - 1
                    
                    target = target + 1
                    weekTarget = weekTarget + 1
                    
                End If
            Next r
        End With
'TO THIS
End Sub
'DID THIS AND TIDIED UP THE PIVOT TABLE
Sub RefreshPivot()
If Worksheets("Daily").Range("C2") > "" Then
    Worksheets("Needs").PivotTables("Day").RefreshTable
    Worksheets("Needs").PivotTables("Day").PivotFields("ship name").ClearAllFilters
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
    printRng.PrintOut
    
    Worksheets("Label").Activate
    
    
End Sub
'DID THIS and in WORKBOOK CODE
Sub ClearDaily()
Dim lastToday As Integer, lastShip As Integer, ships As Range, arr() As Variant, s As Variant
Dim ship As String, path As String

lastToday = Worksheets("Daily").Range("A" & Rows.Count).End(xlUp).Row
lastShip = Worksheets("Daily").Range("F" & Rows.Count).End(xlUp).Row
path = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\"

Set ships = Worksheets("Daily").Range("F2:F" & lastShip)

'arr = ships
    
For Each s In ships
    DeleteDirectory CStr(s)
Next

Worksheets("Daily").Range("A2:F3000").ClearContents


End Sub


Sub ClearFilters()

    If Worksheets("Daily").FilterMode = True Then
        Worksheets("Daily").ShowAllData
    End If

End Sub

Sub DeleteDirectory(ship As String)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim directoryPath As String
    
    path = "C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\"
    
    '"C:\Users\vince\OneDrive\Desktop\Delaware Ship\OrderPDFs\MV BALTIC MANTIS-329022"
    directoryPath = path & ship
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the directory exists
    If fso.FolderExists(directoryPath) Then
        ' Delete files in the directory
        For Each file In fso.GetFolder(directoryPath).Files
            file.Delete
        Next file
        
        ' Delete subdirectories
        For Each folder In fso.GetFolder(directoryPath).SubFolders
            DeleteDirectory folder.path
        Next folder
        
        ' Delete the directory itself (if it's empty)
        On Error Resume Next
        fso.DeleteFolder directoryPath, True
        On Error GoTo 0
        
        
        
        ' Check if directory deletion was successful
        'If fso.FolderExists(directoryPath) Then
        '    MsgBox "Failed to delete directory: " & directoryPath, vbExclamation
        'Else
        '    MsgBox "Directory deleted successfully: " & directoryPath, vbInformation
        'End If
    Else
        MsgBox "Folder not found for " & ship & ", which isn't a bad thing. Click the OK button.", vbExclamation
    End If
    

    
    ' Release the FileSystemObject
    Set fso = Nothing
    
    

End Sub



