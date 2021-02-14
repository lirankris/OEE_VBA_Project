VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////Get the total Hours of the box item data set.//////////////////////////////////////////////////////////////////////////
Sub CountHoursBox()
    
    Dim i As Integer
    Dim len_bottels As Integer
    Dim ws As Worksheet
    Dim count As Integer
    
    Set ws = ThisWorkbook.Worksheets("Raw_data_box")
    
    len_box = ws.Range("L9").Value
    
    With ws
    
        For i = 2 To len_box
        
            If Hour(.Range("B" & i).Value) <> Hour(.Range("B" & (i + 1)).Value) Then
            
                count = count + 1
                
            End If
            
        Next i
        
        .Range("J13").Value = count
        
    End With

End Sub
