Attribute VB_Name = "Add_new_Workbook_data"
Option Explicit

Public len_It, len_Bt, len_Pt As Integer
Public I_time, I_Weight, B_time, B_Weight, P_time, P_Weight As Object

Function Add_new_data_to_sheet(FileName, GetBook, lastRow_file, ws1, ws2, ws3)
    
    Dim Object_Type, Object_Weight, Object_TimeStamp As Range

    Set Object_Type = Workbooks(FileName).Sheets(1).Range("A2:A" & lastRow_file) 'Set the length of Weight array Object.
    Set Object_Weight = Workbooks(FileName).Sheets(1).Range("G2:G" & lastRow_file) 'Set the length of target Weight array Object.
    Set Object_TimeStamp = Workbooks(FileName).Sheets(1).Range("H2:H" & lastRow_file) 'Set the length of timestamp array Object.
    
    Set I_time = CreateObject("System.Collections.ArrayList") 'Create New time array Object.
    Set I_Weight = CreateObject("System.Collections.ArrayList") 'Create New Weight array Object.
    Set B_time = CreateObject("System.Collections.ArrayList") 'Create New time array Object.
    Set B_Weight = CreateObject("System.Collections.ArrayList") 'Create New Weight array Object.
    Set P_time = CreateObject("System.Collections.ArrayList") 'Create New time array Object.
    Set P_Weight = CreateObject("System.Collections.ArrayList") 'Create New Weight array Object.
    
    Dim Obj_n, len_It, len_Iw, len_Bt, len_Bw, len_Pt, len_Pw As Integer
    
    Dim DblTimeStampI As Double
    Dim DblTimeStampB As Double
    Dim DblTimeStampP As Double
    
'/////////////////////////////////////////////Separating between different Objects (Item\Box\ShippingPallet) by columns.///////////////////////////////////////////////////
    For Obj_n = 0 To lastRow_file
    
        If Object_Type(Obj_n) = "Item" Then
        
            DblTimeStampI = CDbl(Object_TimeStamp(Obj_n)) 'Create New Weight "item Weight" item as a Double var'.
            I_time.Add DblTimeStampI
            I_Weight.Add Object_Weight(Obj_n)
          
        ElseIf Object_Type(Obj_n) = "Box" Then 'Create New "Box Weight" item as a Double var'.
            
            DblTimeStampB = CDbl(Object_TimeStamp(Obj_n))
            B_time.Add DblTimeStampB
            B_Weight.Add Object_Weight(Obj_n)
        
        ElseIf Object_Type(Obj_n) = "ShippingPallet" Then 'Create New "Pallet Weight" item as a Double var'.
        
            DblTimeStampP = CDbl(Object_TimeStamp(Obj_n))
            P_time.Add DblTimeStampP
            P_Weight.Add Object_Weight(Obj_n)
        
        End If
        
    Next Obj_n
    
'///////////////////////////////////////////////////////////Get the lentgh of every data set.//////////////////////////////////////////////////////////////////////////
    len_It = I_time.count
    len_Iw = I_Weight.count
    
    len_Bt = B_time.count
    len_Bw = B_Weight.count
    
    len_Pt = P_time.count
    len_Pw = P_Weight.count
    
    If len_It = len_Iw Then
        ws1.Range("H3").Value = len_It
    Else
        ws1.Range("H3").Value = len_Bt
        MsgBox "Some data is missing"
    End If
    
    If len_Bt = len_Bw Then
        ws2.Range("H3").Value = len_Bt
    Else
        ws2.Range("H3").Value = len_Bt
        MsgBox "Some data is missing"
    End If
    
    If len_Pt = len_Pw Then
        ws3.Range("H3").Value = len_Pt
    Else
        ws3.Range("H3").Value = len_Pt
        MsgBox "Some data is missing"
    End If
    
    
    Call all_len(len_It, len_Bt, len_Pt) 'insert the workbook the columns.

    ws1.Range("A2").Resize(len_Iw, 1).Value = Application.WorksheetFunction.Transpose(I_Weight.toArray) 'Transpose the values tn I_Weight to an array to fit to the column.
    ws1.Range("B2").Resize(len_It, 1).Value = Application.WorksheetFunction.Transpose(I_time.toArray) 'Transpose the values tn I_time to an array to fit to the column.
    
'////////////////////////////////////////////////////sort all the new columns with the new data by date and time.////////////////////////////////////////////////////////////////////////

    With ws1.Sort
    
        .SortFields.Clear
        .SortFields.Add2 Key:=ws1.Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws1.Range("A1:B" & len_It)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    ws2.Range("A2").Resize(len_Bw, 1).Value = Application.WorksheetFunction.Transpose(B_Weight.toArray)
    ws2.Range("B2").Resize(len_Bt, 1).Value = Application.WorksheetFunction.Transpose(B_time.toArray)
    
    With ws2.Sort
    
        .SortFields.Clear
        .SortFields.Add2 Key:=ws1.Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws1.Range("A1:B" & len_Bt)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    ws3.Range("A2").Resize(len_Pw, 1).Value = Application.WorksheetFunction.Transpose(P_Weight.toArray)
    ws3.Range("B2").Resize(len_Pt, 1).Value = Application.WorksheetFunction.Transpose(P_time.toArray)
    
    With ws3.Sort
    
        .SortFields.Clear
        .SortFields.Add2 Key:=ws1.Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws1.Range("A1:B" & len_Pt)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
End Function
