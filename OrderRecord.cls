VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OrderRecord class

Private orQuantity As Double
Private orOrderMeasurement As String
Private orOrderItem As String
Private orShip As String

' Property Get and Property Let for each Property
Public Property Get Quantity() As Double
    Quantity = orQuantity
End Property
Public Property Let Quantity(ByVal value As Double)
    orQuantity = value
End Property
Public Property Get OrderMeasurement() As String
    OrderMeasurement = orOrderMeasurement
End Property
Public Property Let OrderMeasurement(ByVal value As String)
    orOrderMeasurement = value
End Property
Public Property Get OrderItem() As String
    OrderItem = orOrderItem
End Property
Public Property Let OrderItem(ByVal value As String)
    orOrderItem = value
End Property
Public Property Get ship() As String
    ship = orShip
End Property
Public Property Let ship(ByVal value As String)
    orShip = value
End Property

' Derived properties

'Use measurment to do a lookup into "MasterList" and return the
'clean name associated with that rowset
Public Property Get CleanMeasurement() As String
    Dim msrmntLookupRange As Range
    Set msrmntLookupRange = Worksheets("Master List").Range("F:G")
    On Error Resume Next
        CleanMeasurement = Application.WorksheetFunction.VLookup(orOrderMeasurement, msrmntLookupRange, 2, False)
    
    'If there's an error make the clean measurement the same as the order measurement
    '*** TODO: make a userform to insert a new rowset to avoid future issues***
    If Err.Number <> 0 Then
        CleanMeasurement = orOrderMeasurement
    End If
End Property

'Use measurment to do a lookup into "MasterList" and return the
'clean name associated with that rowset
Public Property Get CleanItem() As String
    Dim itemLookupRange As Range
    Set itemLookupRange = Worksheets("Master List").Range("B:C")
    On Error Resume Next
        CleanItem = Application.WorksheetFunction.VLookup(orOrderItem, itemLookupRange, 2, False)
    'If there's an error make the clean item the same as the order item
    '*** TODO: make a userform to insert a new rowset to avoid future issues***
    If Err.Number <> 0 Then
        CleanItem = orOrderItem
    End If
End Property

'Look up item and return the case weight associated with that rowset
Public Property Get ItemCaseWeight() As Double
    On Error Resume Next
        ItemCaseWeight = Application.WorksheetFunction.XLookup(CleanItem, Worksheets("Master List").Range("C:C"), Worksheets("Master List").Range("E:E"))
    
    'If not found use order quantity
    'Also will be handled by item error Userform
    If Err.Number <> 0 Then
        ItemCaseWeight = orQuantity
    End If
End Property

'Simple toString function to use for debugging
Public Function toString() As String
    toString = "{" & Quantity & "," & OrderMeasurement & "," & _
                CleanMeasurement & "," & OrderItem & "," & CleanItem & "," & _
                ItemCaseWeight & "}"
End Function
