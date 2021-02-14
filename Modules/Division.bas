Attribute VB_Name = "Division"
Option Explicit

Private ClcMws As Worksheet
Private ClcWs1 As Worksheet
Private ClcWs2 As Worksheet
Private ClcWs3 As Worksheet

Private Sub Div_42(line_name)

    Set ClcMws = ThisWorkbook.Worksheets("Main")
    Set ClcWs1 = ThisWorkbook.Worksheets("Raw_data_item")
    Set ClcWs2 = ThisWorkbook.Worksheets("Raw_data_box")
    Set ClcWs3 = ThisWorkbook.Worksheets("Raw_data_pallet")
    
    If line_name = "10L" Then
    
        Call A10L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "5L" Then
    
        Call A5L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "1L" Then
    
        Call A1L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    End If
    
End Sub

Private Sub Div_11(line_name)

    Set ClcMws = ThisWorkbook.Worksheets("Main")
    Set ClcWs1 = ThisWorkbook.Worksheets("Raw_data_item")
    Set ClcWs2 = ThisWorkbook.Worksheets("Raw_data_box")
    Set ClcWs3 = ThisWorkbook.Worksheets("Raw_data_pallet")
    
    If line_name = "COSTEC 12" Then
    
        Call COSTEC_12(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "COSTEC 4" Then
    
        Call COSTEC_4(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "KOSME" Then
    
        Call KOSME(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "ROVEMA" Then
    
        Call ROVEMA(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    End If
    
End Sub

Private Sub Div_10(line_name)

    Set ClcMws = ThisWorkbook.Worksheets("Main")
    Set ClcWs1 = ThisWorkbook.Worksheets("Raw_data_item")
    Set ClcWs2 = ThisWorkbook.Worksheets("Raw_data_box")
    Set ClcWs3 = ThisWorkbook.Worksheets("Raw_data_pallet")
    
    If line_name = "5/2" Then
    
        Call five2two(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "2/2" Then
    
        Call two2two(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    ElseIf line_name = "1/1" Then
    
        Call one2one(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
    End If

End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////---- line1A ="10L" , line2A ="5L", line3A ="1L" ----///////////////////
'///////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////---- DA = data manipulation ----///////////////////////////////
'//////////////////////////---- M_site_location = Main site location ----/////////////////
'/////////////////////////---- D_Location = Division Location ----///////////////////////
'////////////////////////---- D_line_name = Division D_line_name ----///////////////////
'///////////////////////---- D_Incharge = Division Incharge ----///////////////////////
'//////////////////////---- C_pace = can pace ----////////////////////////////////////
'/////////////////////---- B_pace = bag pace ----////////////////////////////////////
'////////////////////---- I_sb = Item Substance Type ----///////////////////////////
'///////////////////---- C_p_size = can pack size ----/////////////////////////////
'//////////////////---- C_delta = can delta ----//////////////////////////////////
'/////////////////---- B_p_size = bag pack size ----/////////////////////////////
'////////////////---- B_delta = bag delta ----//////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
    
    Sub A10L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line1A As clsSpecs
        Set line1A = New clsSpecs
        
        '10L_Info
        line1A.M_site_location = "location"
        line1A.D_Location = "Division"
        line1A.D_line_name = "10L"
        line1A.D_Incharge = "Incharge"
            'PACE:
        line1A.C_pace = 60
        line1A.P_box_pace = 4
        line1A.P_pallte_pace = 2
            'Size&sb:
        line1A.C_sb = "Liquid"
        line1A.C_p_size = 10
        line1A.max_w = 12
        
        '10L_DA
            'Main:
        With ClcMws
            .Cells(24, 3).Value = line1A.M_site_location
            .Cells(26, 3).Value = line1A.D_Location
            .Cells(3, 5).Value = line1A.D_line_name
            .Cells(28, 3).Value = line1A.D_Incharge
            .Cells(20, 3).Value = line1A.C_pace
            .Cells(22, 3).Value = line1A.C_sb
            .Cells(22, 4).Value = line1A.C_p_size & " Liter"
            .Cells(4, 9).Value = line1A.C_p_size & " Liter"
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line1A.C_pace
        ClcWs1.Cells(11, 14).Value = line1A.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line1A.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line1A.P_pallte_pace

    End Sub

    Sub A5L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line2A As clsSpecs
        Set line2A = New clsSpecs
        
        '10L_Info
        line2A.M_site_location = "location"
        line2A.D_Location = "Division"
        line2A.D_line_name = "5L"
        line2A.D_Incharge = "Incharge"
            'PACE:
        line2A.C_pace = 60
        line2A.P_box_pace = 4
        line2A.P_pallte_pace = 2
            'Size&sb:
        line2A.C_sb = "Liquid"
        line2A.C_p_size = 5
        line2A.max_w = 6.2
        
        '5L_DA
            'Main:
        With ClcMws
            .Cells(24, 3).Value = line2A.M_site_location
            .Cells(26, 3).Value = line2A.D_Location
            .Cells(3, 5).Value = line2A.D_line_name
            .Cells(28, 3).Value = line2A.D_Incharge
            .Cells(20, 3).Value = line2A.C_pace
            .Cells(22, 3).Value = line2A.C_sb
            .Cells(22, 4).Value = line2A.C_p_size & " Liter"
            .Cells(4, 9).Value = line2A.C_p_size & " Liter"
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line2A.C_pace
        ClcWs1.Cells(11, 14).Value = line2A.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line2A.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line2A.P_pallte_pace

    End Sub

    Sub A1L(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line3A As clsSpecs
        Set line3A = New clsSpecs
        
        '10L_Info
        line3A.M_site_location = "location"
        line3A.D_Location = "Division"
        line3A.D_line_name = "1L"
        line3A.D_Incharge = "Incharge"
            'PACE:
        line3A.C_pace = 60
        line3A.P_box_pace = 4
        line3A.P_pallte_pace = 2
            'Size&sb:
        line3A.C_sb = "Liquid"
        line3A.C_p_size = 1
        line3A.max_w = 1.8
        
        '10L_DA
            'Main:
        With ClcMws
            .Cells(24, 3).Value = line3A.M_site_location
            .Cells(26, 3).Value = line3A.D_Location
            .Cells(3, 5).Value = line3A.D_line_name
            .Cells(28, 3).Value = line3A.D_Incharge
            .Cells(20, 3).Value = line3A.C_pace
            .Cells(22, 3).Value = line3A.C_sb
            .Cells(22, 4).Value = line3A.C_p_size & " Liter"
            .Cells(4, 9).Value = line3A.C_p_size & " Liter"
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line3A.C_pace
        ClcWs1.Cells(11, 14).Value = line3A.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line3A.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line3A.P_pallte_pace

    End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////---- line1 ="COSTEC_12" , line2 ="COSTEC_4", line3 ="KOSME", line4 ="ROVEMA" ----//
'///////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////---- DA = data manipulation ----///////////////////////////////
'//////////////////////////---- M_site_location = Main site location ----/////////////////
'/////////////////////////---- D_Location = Division Location ----///////////////////////
'////////////////////////---- D_line_name = Division D_line_name ----///////////////////
'///////////////////////---- D_Incharge = Division Incharge ----///////////////////////
'//////////////////////---- C_pace = can pace ----////////////////////////////////////
'/////////////////////---- B_pace = bag pace ----////////////////////////////////////
'////////////////////---- I_sb = Item Substance Type ----///////////////////////////
'///////////////////---- C_p_size = can pack size ----/////////////////////////////
'//////////////////---- C_delta = can delta ----//////////////////////////////////
'/////////////////---- B_p_size = bag pack size ----/////////////////////////////
'////////////////---- B_delta = bag delta ----//////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////

    Sub COSTEC_12(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line1 As clsSpecs
        Set line1 = New clsSpecs
        
        'COSTEC_12_Info
        line1.M_site_location = "location"
        line1.D_Location = "Division"
        line1.D_line_name = "COSTEC 12"
        line1.D_Incharge = "Incharge"
            'PACE:
        line1.C_pace = 60
        line1.P_box_pace = 4
        line1.P_pallte_pace = 2
            'Size&sb:
        line1.C_sb = "Liquid"
        line1.C_p_size = 1
        line1.max_w = 1.8
        
        'COSTEC_12_DA
            'Main:
        With ClcMws
            .Cells(24, 3).Value = line1.M_site_location
            .Cells(26, 3).Value = line1.D_Location
            .Cells(3, 5).Value = line1.D_line_name
            .Cells(28, 3).Value = line1.D_Incharge
            .Cells(20, 3).Value = line1.C_pace
            .Cells(22, 3).Value = line1.C_sb
            .Cells(22, 4).Value = line1.C_p_size & " Liter"
            .Cells(4, 9).Value = line1.C_p_size & " Liter"
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line1.C_pace
        ClcWs1.Cells(11, 14).Value = line1.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line1.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line1.P_pallte_pace

    End Sub

    Sub COSTEC_4(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line2 As clsSpecs
        Set line2 = New clsSpecs
        
        'COSTEC_4_Info
        line2.M_site_location = "location"
        line2.D_Location = "Division"
        line2.D_line_name = "COSTEC 4"
        line2.D_Incharge = "Incharge"
              'PACE:
        line2.C_pace = 32
        line2.P_box_pace = 2
        line2.P_pallte_pace = 1
            'Size&sb:
        line2.C_sb = "Liquid"
        line2.C_p_size = 1
        line2.max_w = 1.8
        
        'COSTEC_4_DA
            'Main:
        With ClcMws
        
            .Cells(24, 3).Value = line2.M_site_location
            .Cells(26, 3).Value = line2.D_Location
            .Cells(3, 5).Value = line2.D_line_name
            .Cells(28, 3).Value = line2.D_Incharge
            .Cells(20, 3).Value = line2.C_pace
            .Cells(22, 3).Value = line2.C_sb
            .Cells(22, 4).Value = line2.C_p_size & " Liter"
            .Cells(4, 9).Value = line2.C_p_size & " Liter"
            
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line2.C_pace
        ClcWs1.Cells(11, 14).Value = line2.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line2.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line2.P_pallte_pace
        
    End Sub
    
        Sub KOSME(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line3 As clsSpecs
        Set line3 = New clsSpecs
        
        'KOSME_Info
        line3.M_site_location = "location"
        line3.D_Location = "Division"
        line3.D_line_name = "KOSME"
        line3.D_Incharge = "Incharge"
               'PACE:
        line3.C_pace = 40
        line3.P_box_pace = 2
        line3.P_pallte_pace = 1
            'Size&sb:
        line3.C_sb = "Liquid"
        line3.C_p_size = 5
        line3.max_w = 7
        
        'KOSME_DA
           'Main:
        With ClcMws
        
            .Cells(24, 3).Value = line3.M_site_location
            .Cells(26, 3).Value = line3.D_Location
            .Cells(3, 5).Value = line3.D_line_name
            .Cells(28, 3).Value = line3.D_Incharge
            .Cells(20, 3).Value = line3.C_pace
            .Cells(22, 3).Value = line3.C_sb
            .Cells(22, 4).Value = line3.C_p_size & " Liter"
            .Cells(4, 9).Value = line3.C_p_size & " Liter"
            
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line3.C_pace
        ClcWs1.Cells(11, 14).Value = line3.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line3.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line3.P_pallte_pace
        
        End Sub

        Sub ROVEMA(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line4 As clsSpecs
        Set line4 = New clsSpecs
        
        'ROVEMA_Info
        line4.M_site_location = "location"
        line4.D_Location = "Division"
        line4.D_line_name = "ROVEMA"
        line4.D_Incharge = "Incharge"
            'PACE:
        line4.B_pace = 10
        line4.P_box_pace = 2
        line4.P_pallte_pace = 1
            'Size&sb:
        line4.B_sb = "Powder"
        line4.B_p_size = 1
        line4.max_w = 6
        
        'ROVEMA_DA
           'Main:
        With ClcMws
        
            .Cells(24, 3).Value = line4.M_site_location
            .Cells(26, 3).Value = line4.D_Location
            .Cells(3, 5).Value = line4.D_line_name
            .Cells(28, 3).Value = line4.D_Incharge
            .Cells(20, 3).Value = line4.B_pace
            .Cells(22, 3).Value = line4.B_sb
            .Cells(22, 4).Value = line4.B_p_size & " Kg"
            .Cells(4, 9).Value = line4.B_p_size & " Kg"
        
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line4.B_pace
        ClcWs1.Cells(11, 14).Value = line4.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line4.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line4.P_pallte_pace
        
        
        End Sub
        
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////---- line5 ="5/2" , line6 ="2/2", line7 ="1/1" ----////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////---- DA = data manipulation ----///////////////////////////////
'//////////////////////////---- M_site_location = Main site location ----/////////////////
'/////////////////////////---- D_Location = Division Location ----///////////////////////
'////////////////////////---- D_line_name = Division D_line_name ----///////////////////
'///////////////////////---- D_Incharge = Division Incharge ----///////////////////////
'//////////////////////---- C_pace = can pace ----////////////////////////////////////
'/////////////////////---- B_pace = bag pace ----////////////////////////////////////
'////////////////////---- I_sb = Item Substance Type ----///////////////////////////
'///////////////////---- C_p_size = can pack size ----/////////////////////////////
'//////////////////---- C_delta = can delta ----//////////////////////////////////
'/////////////////---- B_p_size = bag pack size ----/////////////////////////////
'////////////////---- B_delta = bag delta ----//////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////

        Sub five2two(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line5 As clsSpecs
        Set line5 = New clsSpecs
        
        '5/2_Info
        line5.M_site_location = "location"
        line5.D_Location = "Division"
        line5.D_line_name = "5\2"
        line5.D_Incharge = "Incharge"
            'PACE:
        line5.C_pace = 30
        line5.P_box_pace = 3
        line5.P_pallte_pace = 2
            'Size&sb:
        line5.C_sb = "Liquid"
        line5.C_p_size = 5
        line5.max_w = 7
        
        '5/2_DA
           'Main:
        With ClcMws
        
            .Cells(24, 3).Value = line5.M_site_location
            .Cells(26, 3).Value = line5.D_Location
            .Cells(3, 5).Value = line5.D_line_name
            .Cells(28, 3).Value = line5.D_Incharge
            .Cells(20, 3).Value = line5.C_pace
            .Cells(22, 3).Value = line5.C_sb
            .Cells(22, 4).Value = line5.C_p_size & " Liter"
            .Cells(4, 9).Value = line5.C_p_size & " Liter"
            
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line5.C_pace
        ClcWs1.Cells(11, 14).Value = line5.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line5.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line5.P_pallte_pace
        
        End Sub
        
        Sub two2two(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line6 As clsSpecs
        Set line6 = New clsSpecs
        
        '2/2_Info
        line6.M_site_location = "location"
        line6.D_Location = "Division"
        line6.D_line_name = "2\2"
        line6.D_Incharge = "Incharge"
           'PACE:
        line6.C_pace = 30
        line6.P_box_pace = 2
        line6.P_pallte_pace = 1
            'Size&sb:
        line6.C_sb = "Liquid"
        line6.C_p_size = 5
        line6.max_w = 7
        
        '2/2_DA
            'Main:
        With ClcMws
        
            .Cells(24, 3).Value = line6.M_site_location
            .Cells(26, 3).Value = line6.D_Location
            .Cells(3, 5).Value = line6.D_line_name
            .Cells(28, 3).Value = line6.D_Incharge
            .Cells(20, 3).Value = line6.C_pace
            .Cells(22, 3).Value = line6.C_sb
            .Cells(22, 4).Value = line6.C_p_size & " Liter"
            .Cells(4, 9).Value = line6.C_p_size & " Liter"
            
        End With
        
            'Bottles:
        ClcWs1.Cells(9, 10).Value = line6.C_pace
        ClcWs1.Cells(11, 14).Value = line6.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line6.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line6.P_pallte_pace
        
        End Sub
        
        Sub one2one(ClcMws, ClcWs1, ClcWs2, ClcWs3)
        
        Dim line7 As clsSpecs
        Set line7 = New clsSpecs
        
        '1/1_Info
        line7.M_site_location = "location"
        line7.D_Location = "Division"
        line7.D_line_name = "1\1"
        line7.D_Incharge = "Incharge"
           'PACE:
        line7.C_pace = 10
        line7.P_box_pace = 1
        line7.P_pallte_pace = 1
            'Size&sb:
        line7.C_sb = "Liquid"
        line7.C_p_size = 10
        line6.max_w = 12
        
        'If "" Then
        
           'line7.C_p_size = 10
            
        'ElseIf "" Then
        
            'line7.C_p_size = 20
            
        'End If "
        
        '1/1_DA
            'Main:
          With ClcMws
        
            .Cells(24, 3).Value = line7.M_site_location
            .Cells(26, 3).Value = line7.D_Location
            .Cells(3, 5).Value = line7.D_line_name
            .Cells(28, 3).Value = line7.D_Incharge
            .Cells(20, 3).Value = line7.C_pace
            .Cells(22, 3).Value = line7.C_sb
            .Cells(22, 4).Value = line7.C_p_size & " Liter"
            .Cells(4, 9).Value = line7.C_p_size & " Liter"
            
        End With

            'Bottles:
        ClcWs1.Cells(9, 10).Value = line7.C_pace
        ClcWs1.Cells(11, 14).Value = line7.max_w
            'Box:
        ClcWs2.Cells(9, 10).Value = line7.P_box_pace
            'Pallet:
        ClcWs3.Cells(9, 10).Value = line7.P_pallte_pace
        
        End Sub
