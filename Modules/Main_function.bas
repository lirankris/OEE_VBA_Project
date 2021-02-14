Attribute VB_Name = "main_function"
Public ws1 As Worksheet
Public ws2 As Worksheet
Public ws3 As Worksheet

Public len1 As Integer
Public len2 As Integer
Public len3 As Integer

Public exception As String

Sub Main()
     
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    
    Dim fso As New FileSystemObject
    Dim FileName As String
    
    Dim Object_Type As Range
    Dim Object_Weight As Range
    Dim Object_TimeStamp As Range
    
    Application.ScreenUpdating = False 'Cancel Screen Updating will Main() run.
    Application.DisplayAlerts = False 'Cancel  Display Alerts will Main() run.
    
'/////////////////////////////////////////////////////////////Open file & get files names////////////////////////////////////////////////////////////////////

    'Opens a dialog box for selecting file from clicking insert file macro.
    FileToOpen = Application.GetOpenFilename(FileFilter:="Excel Files(*.xlsx*),*xlsx*", Title:="Browse for your File & Import Range")
    'GetOpenFilename(
            'FileFilter:="Excel Files(*.xlsx*),*xlsx*" - file filtering criteria to show only xlsx files.
            'FilterIndex - None.
            'Title:="Browse for your File & Import Range" - Specifies the title of the dialog box.
            'ButtonText - None.
            'MultiSelect - None.
            ')

    'If file wasn't selected then close sub.
    If FileToOpen = False Then
        Exit Sub
    End If
    
    'Files names:
    FileName = fso.GetFileName(FileToOpen) 'GetFileName method returns selected file name with (.xlsx).
    BatchNum = Replace(FileName, ".xlsx", "") 'Get File Name (Batch num) witout (.xlsx)
    GetBook = ActiveWorkbook.Name 'GetFileName method returns selected file name.
    
'/////////////////////////////////////////////////////////////Sheet to use////////////////////////////////////////////////////////////////////

    '***When deploy change to: Set ws(i) = Workbooks(GetBook).Worksheets("Raw_data_item\box\pallet").Unprotect "Password"***
    Set ws1 = Workbooks(GetBook).Worksheets("Raw_data_item") 'Get the active Workbook thats in use, then selecr Worksheet "Raw_data_item".
    Set ws2 = Workbooks(GetBook).Worksheets("Raw_data_box") ''Get the active Workbook thats in use, then selecr Worksheet "Raw_data_box".
    Set ws3 = Workbooks(GetBook).Worksheets("Raw_data_pallet") ''Get the active Workbook thats in use, then selecr Worksheet "Raw_data_pallet".
    
'/////////////////////////////////////////////////////Get Data and insert to Workbook/////////////////////////////////////////////////////////
    
    If FileToOpen <> False Then 'If GetOpenFilename method returns selected file True (A file was selected) continue.
    
        Set OpenBook = Workbooks.Open(FileToOpen)
        
        '***A pivot table is depended on the data, clear A1 or B1 may delete the conection to the pivot.***
        ws1.Range("A2:A10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        ws1.Range("B2:B10000").ClearContents 'Clear content from column B, sheet "Raw_data_item" (to enter new data)
        ws1.Range("D2:D10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        ws1.Range("E2:E10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        ws1.Range("G2:G10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        
        ws2.Range("A2:A10000").ClearContents 'Clear content from column A, sheet "Raw_data_box" (to enter new data)
        ws2.Range("B2:B10000").ClearContents 'Clear content from column B, sheet "Raw_data_box" (to enter new data)
        ws2.Range("D2:D10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        ws2.Range("E2:E10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        
        ws3.Range("A2:A10000").ClearContents 'Clear content from column A, sheet "Raw_data_pallet" (to enter new data)
        ws3.Range("B2:B10000").ClearContents 'Clear content from column B, sheet "Raw_data_pallet" (to enter new data)
        ws3.Range("D2:D10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        ws3.Range("E2:E10000").ClearContents 'Clear content from column A, sheet "Raw_data_item" (to enter new data)
        
        lastRow_file = Workbooks(FileName).Worksheets(1).Cells(Worksheets(1).Rows.count, "A").End(xlUp).Row 'Get the last row of the new data file (FileToOpen).
        
        'Returns a Range objects from row 2 to lastRow (Not blank) the new data file.
        
        Call Add_new_data_to_sheet(FileName, GetBook, lastRow_file, ws1, ws2, ws3) 'Insert new data to ThisWorkbook.
        
        Application.CutCopyMode = False
        OpenBook.Close False 'Closing the new data workbook.
           
    End If
    
    'On Error Resume Next 'Ignores the error and continues on.
        
        line_drop_list.Show vbModal 'Call the interactiv chat box.
        
    'On Error GoTo 0 'When error occurs, the code stops and displays the error.
        
            'MsgBox ("sorry something want wrong")
     
        Call Delta_Item(ws1, len1) 'Get the delta of the item.
        Call Delta_Box(ws2, len2) 'Get the delta of the Box.
        Call Delta_Pallet(ws3, len3) 'Get the delta of the Pallet.
        
        Call scroll_bar_Item(ws1, len1) 'insert data to view table item (Weight).
        Call scroll_bar_Box(ws2, len2) 'insert data to view table Box (Weight).
        Call scroll_bar_Pallet(ws3, len3) 'insert data to view table Pallet (Weight).
      
        Application.Run "Sheet4.ScrollBar_Change_bottles", ws1 'Set the max value of the scroll bar.
        Application.Run "Sheet11.ScrollBar_Change_box", ws2 'Set the max value of the scroll bar.
        Application.Run "Sheet8.ScrollBar_Change_pallte", ws3 'Set the max value of the scroll bar.
        
        Call clear_array
        
        ActiveWorkbook.RefreshAll
        
        'BatchNum = InputBox("Enter Batch Number:") 'get the Batch Number from the user.
        TargetVal = InputBox("Enter Target weight:") 'get the Target value from the user.
        
        With ThisWorkbook.Worksheets("Main")
        
            .Cells(14, 3).Value = BatchNum
            .Cells(10, 3).Value = TargetVal
            
        End With
        
        ws1.Cells(2, 6).Value = BatchNum
        ws1.Cells(9, 14).Value = TargetVal
        
        Application.Run "Sheet6.Group_pivot" 'Group All pivot table.
     
        Application.Run "Sheet5.CountHoursBottels"
        Application.Run "Sheet9.CountHoursBox"
        Application.Run "Sheet10.CountHoursPallet"
        
        Application.DisplayAlerts = True 'Activate Display Alerts.
        Application.ScreenUpdating = True 'Activate Screen Updating.
        
        Application.DisplayFullScreen = True
        
        MsgBox "Made by liran krispin ;)"
    
End Sub

Public Sub all_len(len_It, len_Bt, len_Pt)
        
        len1 = len_It
        len2 = len_Bt
        len3 = len_Pt
            
End Sub
