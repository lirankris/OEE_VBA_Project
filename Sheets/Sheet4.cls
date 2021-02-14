VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////set the scrollbar lentgh Value to the bottles count max.//////////////////////////////////////////////////////////////////////////
Private Sub ScrollBar_Change_bottles(ws1)

    Dim sh As Worksheet
    Dim Bar As ScrollBar
    
    For Each sh In ThisWorkbook.Worksheets
    
        For Each Bar In sh.ScrollBars
        
            If Bar.Name = "Scroll Bar 3" Or Bar.Name = "Scroll Bar 4" Then
            
                Max_Value = ws1.Range("H3").Value
                Bar.Max = Max_Value + 1
           
            End If
        Next Bar
    Next sh
End Sub
