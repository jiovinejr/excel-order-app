VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Event code, Target variable contains the cell or range that has been changed
Private Sub Worksheet_Change(ByVal target As Range)
 
'Interesect method returns a range object of a rectangular intersection of two or more cell ranges
If Not Intersect(target, Range("E1")) Is Nothing Then
    On Error Resume Next
        GetSticksForOrder
        
End If
End Sub

