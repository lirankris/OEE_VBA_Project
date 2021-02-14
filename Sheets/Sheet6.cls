VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////Group all pivots.//////////////////////////////////////////////////////////////////////////
Private Sub Group_pivot()

    Worksheets("OVER_deltaT").Visible = True
    
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim ptf As PivotField
    Dim pti As PivotItem
    Dim MyRange As Range
    
    Set ws = ThisWorkbook.Worksheets("OVER_deltaT")
    Set MyRange = ws.Range("A1")
    
    With ws
        For Each pt In .PivotTables
            For Each ptf In pt.RowFields
                    If ptf.Name = "Date created" Then
                        ptf.DrillTo ptf.Name
                    End If
            Next ptf
        Next pt
    End With
    
    Worksheets("OVER_deltaT").Visible = False
    Worksheets("Sum_of_Weight").Visible = False
    
    
End Sub

