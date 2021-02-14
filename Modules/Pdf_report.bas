Attribute VB_Name = "Pdf_report"
Option Explicit

Public OutputFolder As String

Sub createPdf()

    Dim FileName As String
    Dim ws As Worksheet
    Dim id As Dictionary
    Dim user_input As String
    Dim result As Boolean
    
    Set ws = ThisWorkbook.Worksheets("Pdf")
    Set id = CreateObject("Scripting.Dictionary")
    
    id.Add user_input, "Liran Krispin"
    
    If id.Exists(user_input) Then
    
        ws.Visible = True
    
        With ws
            
            result = SavePdfile()
            
            If result = False Then
                Exit Sub
                
            Else
                .Range("C48").Value = "John Doe"
                
                FileName = .Range("C2").text & "_QA Weight"
                
                .ExportAsFixedFormat _
                Type:=xlTypePDF, _
                FileName:=OutputFolder + "\" + FileName + ".pdf", _
                Ignoreprintareas:=False, _
                openafterpublish:=True
                
            End If
        End With
    Else
        MsgBox "Wrong ID"
        Exit Sub
    End If

    ws.Visible = False
End Sub

Function SavePdfile() As Boolean

    Application.DisplayAlerts = False
    
    Dim OpenExcelDialogBox As Object
    Dim SaveDialogBox As Object
    Dim MySelectedFile As String

    'Bring the macro file in front
    Windows(ThisWorkbook.Name).Visible = True
    AppActivate Application.Caption
    
    'Select output folder where output files will be saved
    Set SaveDialogBox = Application.FileDialog(msoFileDialogFolderPicker)
    
    If SaveDialogBox.Show = False Then
        SavePdfile = False
        Exit Function
    Else
        SavePdfile = True
        OutputFolder = SaveDialogBox.SelectedItems(1)
    End If

End Function
