Attribute VB_Name = "Module3"
Sub PrintBoxLabels(begin As Variant, last As Variant)

Dim sheetName As String
Dim ObjDoc As bpac.Document, kg As Double
    Set ObjDoc = CreateObject("bpac.Document")

    labelPath = "LABEL FILE PATH GOES HERE"
    shipName = Worksheets("Label").Range("E1").Text
    
  
    sheetName = ActiveSheet.Name
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoCutAtEnd
            
            For i = begin To last
                kg = Format(Range("A" & i) / 2.2, "0.00")
                qty = Range("A" & i).Text
                meas = Range("B" & i).Text
                item = Range("C" & i).Text
                
                If sheetName = "Daily" Then
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

Sub PrintSkidLabel()

    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")
    
    labelPath = "LABEL FILE PATH GOES HERE"
    shipName = Worksheets("Label").Range("E1").Text
    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("ShipName").Text = shipName
            ObjDoc.PrintOut 1, bpoDefault
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close

    
End Sub

Sub PrintRollLabel(ship As String)

    Dim ObjDoc As bpac.Document
    Set ObjDoc = CreateObject("bpac.Document")
    
    labelPath = "LABEL FILE PATH GOES HERE"
    shipSend = ship

    
    ObjDoc.Open (labelPath)
        ObjDoc.StartPrint "", bpoDefault
            ObjDoc.GetObject("RollLabel").Text = shipSend
            ObjDoc.PrintOut 1, bpoDefault
        ObjDoc.EndPrint
    ObjDoc.Close
End Sub

Sub MultiSkid()

Dim ObjDoc As bpac.Document
Set ObjDoc = CreateObject("bpac.Document")

labelPath = "LABEL FILE PATH GOES HERE"
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

