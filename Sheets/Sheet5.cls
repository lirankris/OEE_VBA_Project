VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////Get the total Hours of the Bottels item data set.//////////////////////////////////////////////////////////////////////////
Sub CountHoursBottels()
    
    Dim i As Integer
    Dim len_bottels As Range
    Dim ws As Worksheet
    Dim count As Integer
    Dim PreHourSumWeight As Integer
    
    Set ws = ThisWorkbook.Worksheets("Raw_data_item")
    Set len_bottels = ws.Range("L9")
    
    With ws
        
        PreHourSumWeight = 60 * (.Range("J9").Value) * 0.4 'predicte
        HourToFinish = .Range("J15").Value
        RemHour = .Range("J16").Value
        
        For i = 2 To len_bottels.Value
        
            If Hour(.Range("B" & i).Value) <> Hour(.Range("B" & (i + 1)).Value) Then
            
                count = count + 1
                
                If count <= HourToFinish Then
                
                    .Range("G" & i).Value = PreHourSumWeight
                    
                ElseIf count > HourToFinish And (count - HourToFinish) < 1 Then
                
                    .Range("G" & i).Value = PreHourSumWeight * RemHour
                
                End If
            End If
        Next i
        
        .Range("J13").Value = count
        
    End With
End Sub
