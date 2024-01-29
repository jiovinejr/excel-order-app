VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9765.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBox1_Click()

End Sub

Private Sub Ok_Click()

Dim answer As Integer

answer = MsgBox("Is today a new day?", vbYesNo, "Daily Needs")

If answer = vbYes Then
    ClearDaily
    Add_To_Daily
Else
    Add_To_Daily
End If

End Sub

Private Sub UserForm_Initialize()
Dim rng As Variant
Dim l As Integer

l = Worksheets("On Deck").Range("F" & Rows.Count).End(xlUp).Row
Set rng = Worksheets("On Deck").Range("F1:F" & l)

With ListBox1
    .RowSource = "'On Deck'!F2:F" & l
End With
End Sub

Sub Add_To_Daily()

UserForm1.Hide

Dim ck As Integer, i As Integer, target As Integer

ck = 0
target = Worksheets("Daily").Range("F" & Rows.Count).End(xlUp).Row + 1

For i = 0 To Me.ListBox1.ListCount - 1
    If Me.ListBox1.Selected(i) Then
        ck = 1
        Worksheets("Daily").Range("F" & target) = Me.ListBox1.List(i)
        target = target + 1
    End If
Next i

If ck = 0 Then
    MsgBox "Nothing Selected"
End If

GetTodaysList
Worksheets("Needs").Activate

End Sub
