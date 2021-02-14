VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} line_drop_list 
   Caption         =   "Line Name"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "line_drop_list.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "line_drop_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Insert_Click()

End Sub

Private Sub UserForm_Initialize()
    
    With ComboBox1
        .AddItem "COSTEC 12"
        .AddItem "COSTEC 4"
        .AddItem "KOSME"
        .AddItem "ROVEMA"
        .AddItem "5/2"
        .AddItem "2/2"
        .AddItem "1/1"
        .AddItem "10L"
        .AddItem "5L"
        .AddItem "1L"
    End With
    
End Sub

Private Sub ComboBox1_Change()
    
    Unload line_drop_list
    
    Set sh = ThisWorkbook.Worksheets("Main")
    
    line_name = Me.ComboBox1.Value
    
    If line_name = "KOSME" Or line_name = "COSTEC 4" Or line_name = "COSTEC 12" Or line_name = "ROVEMA" Then
    
        Application.Run "Div_11", line_name
    
    ElseIf line_name = "5/2" Or line_name = "2/2" Or line_name = "1/1" Then
    
        Application.Run "Div_10", line_name
        
    ElseIf line_name = "10L" Or line_name = "5L" Or line_name = "1L" Then
    
        Application.Run "Div_42", line_name

    End If
    
End Sub





    

