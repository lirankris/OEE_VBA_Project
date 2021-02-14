VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////Set the scrollbar lentgh Value to the pallet count max.//////////////////////////////////////////////////////////////////////////
Private Sub ScrollBar_Change_pallte(ws3)

    Dim sh As Worksheet
    Dim Bar As ScrollBar
    
    For Each sh In ThisWorkbook.Worksheets
    
        For Each Bar In sh.ScrollBars
        
            If Bar.Name = "Scroll Bar 7" Or Bar.Name = "Scroll Bar 8" Then
            
                Max_Value = ws3.Range("H3").Value
                Bar.Max = Max_Value + 1
           
            End If
        Next Bar
    Next sh
End Sub
