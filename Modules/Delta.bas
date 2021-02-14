Attribute VB_Name = "delta"
Private i As Integer
Private I1 As Integer
Private I2 As Integer
Private DTimeIn As String
Private DTimeOut As String



Sub Delta_Item(ws1, len1)
    
    i = 0
    I1 = 0
    I2 = 0
    
    For i = 3 To len1 'The first row is the headline and the second is the first can (can't have delta from nothing..)
    
        With ws1 'ws1 = worksheet(Raw_data_item)
            
            If Hour(.Range("B2")) >= 7 And Hour(.Range("B2")) < 15 Then
            
                .Range("AH2").Value = "Morning"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the morning shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
                
                
            ElseIf Hour(.Range("B2")) >= 15 And Hour(.Range("B2")) < 20 Then
            
                .Range("AH2").Value = "After noon"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the After noon shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            ElseIf Hour(.Range("B2")) >= 22 Or Hour(.Range("B2")) < 7 Then
            
                .Range("AH2").Value = "Night"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the Night shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            End If
            
            I1 = i + 1
            I2 = i - 1
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Morning shift ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) < Hour(i) && (Timestamp(i+1) > (Timestamp(i)+) && Hour(i+1) > 06:00:00 ----/////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 07:00:00 < 21:00:00 And (27/01/2021 07:00:00 = 44223.291) > (26/01/2021 21:00:00 = 44222.875) And 07:00:00 > 06:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            If Hour(.Range("B" & I1)) < Hour(.Range("B" & i)) And .Range("B" & I1) > .Range("B" & i) And Hour(.Range("B" & I1)) > 6 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "Morning"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the morning shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len1).Value 'Insert cell AL3 the Morning shift ended value value if the condition is met.
                        
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- After noon shift ----//////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 14:00:00 && Time(Timestamp(LastRow)) =< 20:00:00  ----//
'//////////////////////////////////////////////////////---- example: 16:00:00 > 15:00:00 And 15:00:00 > 14:00:00 And 15:00:00 < 20:00:00 ----/////////
'//////////////////////////////////////////////////////---- extra: from the 15:00:00 to 20:00:00!. ----//////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & I1)) > 14 And (.Range("B" & len1) Mod 1) <= 0.83333 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the after noon shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len1).Value 'Insert cell AL3 the after noon shift ended value value if the condition is met.
                
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Night shift ----/////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(LastRow) < 07:00:00 && (Time(Timestamp(i+1)) >= 21:00:00 || Time(Timestamp(i+1)) =< 7:00:00 )----////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 15:00:00 And 05:00:00 < 07:00:00 And (21:00:00 >= 21:00:00 or 21:00:00 <= 07:00:00 ) ----///////////////////////////
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----///////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & len1)) < 7 And ((.Range("B" & I1) Mod 1) >= 0.875 Or (.Range("B" & I1) Mod 1) <= 0.291666667) Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the night shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len1).Value 'Insert cell AL3 the after night shift ended value value if the condition is met.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Lunch ----//////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 20:00:00 ----////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 15:00:00 And 21:00:00 > 20:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----//////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf (Hour(.Range("B" & i)) = 11 Or Hour(.Range("B" & i)) = 12) And (.Range("B" & I1) - .Range("B" & i)) > "0:20:00" Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AN2").Value = "Lunch" 'Insert cell AN2 the shift number if the condition is met.
                .Range("AQ2").Value = .Range("B" & I1).Value 'Insert cell AJ3 the night shift start value if the condition is met.
                .Range("AO2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////---- Row is Empty ----////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf IsEmpty(.Range("A" & i).Value) = True Then
            
                .Range("E" & i).Value = ""
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) > Hour(i-1) && Hour(i) = Hour(i+1) ----//////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 20:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            ElseIf Hour(.Range("B" & i)) > Hour(.Range("B" & I2)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) Then
            
                .Range("E" & i).Formula = "=IF((B" & i & "- FLOOR(B" & i & ", ""1:00""))>J11,((B" & i & "- FLOOR(B" & i & ", ""1:00""))-J11),"""")" 'Insert cell E(i) the diff' between Timestamp and Roundown (01:00:00) minus the delta
                                                                                                                                                    'if the diff' between Timestamp and Roundown is bigger then delta.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) = (Hour(i+1)-1) && Hour(i) = Hour(i-1) ----//////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 = 21:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////---- extra: Hour(i) = (Hour(i+1)-1) - difference no more then one hour. ----////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) - 1 And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
                
                .Range("E" & i).Formula = "=IF(((Ceiling(B" & i & ",""1:00"")-B" & i & _
                    ")+(B" & i & "-B" & I2 & "))>J11,((Ceiling(B" & i & ", ""1:00"")-B" & i & ")+(B" & i & " - B" & I2 & ")-J11),"""")" 'Insert cell E(i) [ the diff' between Roundup (02:00:00) and Timestamp(i) plus the diff' between Timestamp(i) and Timestamp(i-1) minus the delta ]
                                                                                                                                        'if [ the diff' between Roundup and Timestamp(i) plus the diff of Timestamp(i) and Timestamp(i-1)] is bigger then delta.
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- (Timestamp(i) - Timestamp(i-1)) > Delta && Hour(i+1) = Hour(i) && Hour(i) = Hour(i-1)----//////////////////////////////
'//////////////////////////////////////////////////////---- example: (27/01/2021 07:03:00)-(27/01/2021 07:02:00) > 00:00:15 And 07:00:00 = 07:00:00 And 07:00:00 = 07:00:00----///
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----/////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

             ElseIf .Range("B" & i) - .Range("B" & I2) > .Range("J11") And Hour(.Range("B" & I1)) = Hour(.Range("B" & i)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
            
                .Range("E" & i).Value = (.Range("B" & i).Value - .Range("B" & I2).Value) - .Range("J11").Value 'Insert cell E(i) [the diff' between Timestamp(i) minus Timestamp(i-1) minus the delta ].
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- No delta needed ----/////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Else
            
                .Range("E" & i).Value = ""
            
            End If
                
        End With
        
    Next i
    
End Sub

Sub Delta_Box(ws2, len2)

    i = 0
    I1 = 0
    I2 = 0
    
    For i = 3 To len2 'The first row is the headline and the second is the first box (can't have delta from nothing..)
    
        With ws2 'ws2 = worksheet(Raw_data_box)
            
            If Hour(.Range("B2")) >= 7 And Hour(.Range("B2")) < 15 Then
            
                .Range("AH2").Value = "Morning"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the morning shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
                
                
            ElseIf Hour(.Range("B2")) >= 15 And Hour(.Range("B2")) < 20 Then
            
                .Range("AH2").Value = "After noon"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the After noon shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            ElseIf Hour(.Range("B2")) >= 22 Or Hour(.Range("B2")) < 7 Then
            
                .Range("AH2").Value = "Night"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the Night shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            End If
            
            I1 = i + 1
            I2 = i - 1
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Morning shift ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) < Hour(i) && (Timestamp(i+1) > (Timestamp(i)+) && Hour(i+1) > 06:00:00 ----/////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 07:00:00 < 21:00:00 And (27/01/2021 07:00:00 = 44223.291) > (26/01/2021 21:00:00 = 44222.875) And 07:00:00 > 06:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            If Hour(.Range("B" & I1)) < Hour(.Range("B" & i)) And .Range("B" & I1) > .Range("B" & i) And Hour(.Range("B" & I1)) > 6 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "Morning"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the morning shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len2).Value 'Insert cell AL3 the Morning shift ended value value if the condition is met.
                
          
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- After noon shift ----///////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 14:00:00 && Hour(i+1) =< 20:00:00 ----///////////
'//////////////////////////////////////////////////////---- example: 16:00:00 > 15:00:00 And 15:00:00 > 14:00:00 And 15:00:00 < 20:00:00 ----//
'//////////////////////////////////////////////////////---- extra: from the 15:00:00 to 20:00:00!. ----///////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & I1)) > 14 And Hour(.Range("B" & I1)) <= 20 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the after noon shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len2).Value 'Insert cell AL3 the after noon shift ended value value if the condition is met.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Night shift ----////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 20:00:00 ----////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 15:00:00 And 21:00:00 > 20:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----//////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & I1)) > 20 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the night shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len2).Value 'Insert cell AL3 the after night shift ended value value if the condition is met.

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////---- Row is Empty ----////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf IsEmpty(.Range("A" & i).Value) = True Then
            
                .Range("E" & i).Value = ""
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) > Hour(i-1) && Hour(i) = Hour(i+1) ----//////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 20:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            ElseIf Hour(.Range("B" & i)) > Hour(.Range("B" & I2)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) Then
            
                .Range("E" & i).Formula = "=IF((B" & i & "- FLOOR(B" & i & ", ""1:00""))>J11,((B" & i & "- FLOOR(B" & i & ", ""1:00""))-J11),"""")" 'Insert cell E(i) the diff' between Timestamp and Roundown (01:00:00) minus the delta
                                                                                                                                                    'if the diff' between Timestamp and Roundown is bigger then delta.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) = (Hour(i+1)-1) && Hour(i) = Hour(i-1) ----//////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 = 21:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////---- extra: Hour(i) = (Hour(i+1)-1) - difference no more then one hour. ----////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) - 1 And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
                
                .Range("E" & i).Formula = "=IF(((Ceiling(B" & i & ",""1:00"")-B" & i & _
                    ")+(B" & i & "-B" & I2 & "))>J11,((Ceiling(B" & i & ", ""1:00"")-B" & i & ")+(B" & i & " - B" & I2 & ")-J11),"""")" 'Insert cell E(i) [ the diff' between Roundup (02:00:00) and Timestamp(i) plus the diff' between Timestamp(i) and Timestamp(i-1) minus the delta ]
                                                                                                                                        'if [ the diff' between Roundup and Timestamp(i) plus the diff of Timestamp(i) and Timestamp(i-1)] is bigger then delta.
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- (Timestamp(i) - Timestamp(i-1)) > Delta && Hour(i+1) = Hour(i) && Hour(i) = Hour(i-1)----//////////////////////////////
'//////////////////////////////////////////////////////---- example: (27/01/2021 07:03:00)-(27/01/2021 07:02:00) > 00:00:15 And 07:00:00 = 07:00:00 And 07:00:00 = 07:00:00----///
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----/////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

             ElseIf .Range("B" & i) - .Range("B" & I2) > .Range("J11") And Hour(.Range("B" & I1)) = Hour(.Range("B" & i)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
            
                .Range("E" & i).Value = (.Range("B" & i).Value - .Range("B" & I2).Value) - .Range("J11").Value 'Insert cell E(i) [the diff' between Timestamp(i) minus Timestamp(i-1) minus the delta ].
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- No delta needed ----/////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Else
            
                .Range("E" & i).Value = ""
            
            End If
                
        End With
        
    Next i
    
End Sub
Sub Delta_Pallet(ws3, len3)
    
    i = 0
    I1 = 0
    I2 = 0
    
    For i = 3 To len3 'The first row is the headline and the second is the first pallet (can't have delta from nothing..)
    
        With ws3 'ws3 = worksheet(Raw_data_pallet)
        
            If Hour(.Range("B2")) >= 7 And Hour(.Range("B2")) < 15 Then
            
                .Range("AH2").Value = "Morning"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the morning shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
                
                
            ElseIf Hour(.Range("B2")) >= 15 And Hour(.Range("B2")) < 20 Then
            
                .Range("AH2").Value = "After noon"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the After noon shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            ElseIf Hour(.Range("B2")) >= 22 Or Hour(.Range("B2")) < 7 Then
            
                .Range("AH2").Value = "Night"
                .Range("AJ2").Value = .Range("B2").Value 'Insert cell AJ2 the Night shift start value if the condition is met.
                .Range("AG2").Value = "shift 1" 'Insert cell AG2 the shift number if the condition is met.
            
            End If
            
            I1 = i + 1
            I2 = i - 1
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Morning shift ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) < Hour(i) && (Timestamp(i+1) > (Timestamp(i)+) && Hour(i+1) > 06:00:00 ----/////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 07:00:00 < 21:00:00 And (27/01/2021 07:00:00 = 44223.291) > (26/01/2021 21:00:00 = 44222.875) And 07:00:00 > 06:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            If Hour(.Range("B" & I1)) < Hour(.Range("B" & i)) And .Range("B" & I1) > .Range("B" & i) And Hour(.Range("B" & I1)) > 6 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "Morning"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the morning shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len3).Value 'Insert cell AL3 the Morning shift ended value value if the condition is met.
                
          
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- After noon shift ----///////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 14:00:00 && Hour(i+1) =< 20:00:00 ----///////////
'//////////////////////////////////////////////////////---- example: 16:00:00 > 15:00:00 And 15:00:00 > 14:00:00 And 15:00:00 < 20:00:00 ----//
'//////////////////////////////////////////////////////---- extra: from the 15:00:00 to 20:00:00!. ----///////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & I1)) > 14 And Hour(.Range("B" & I1)) <= 20 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the after noon shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len3).Value 'Insert cell AL3 the after noon shift ended value value if the condition is met.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////--- Night shift ----////////////////////////
'//////////////////////////////////////////////////////---- Hour(i+1) > Hour(i) && Hour(i+1) > 20:00:00 ----////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 15:00:00 And 21:00:00 > 20:00:00 ----//
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----//////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & I1)) > Hour(.Range("B" & i)) And Hour(.Range("B" & I1)) > 20 Then
            
                .Range("E" & i).Value = "" 'Insert cell E(i) empty value.
                .Range("AG3").Value = "shift 2" 'Insert cell AG3 the shift number if the condition is met.
                .Range("AH3").Value = "After noon"
                .Range("AJ3").Value = .Range("B" & I1).Value 'Insert cell AJ3 the night shift start value if the condition is met.
                .Range("AL2").Value = .Range("B" & i).Value 'Insert cell AL2 when the shift before ended value value if the condition is met.
                .Range("AL3").Value = .Range("B" & len3).Value 'Insert cell AL3 the after night shift ended value value if the condition is met.

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////---- Row is Empty ----////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf IsEmpty(.Range("A" & i).Value) = True Then
            
                .Range("E" & i).Value = ""
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) > Hour(i-1) && Hour(i) = Hour(i+1) ----//////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 > 20:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            ElseIf Hour(.Range("B" & i)) > Hour(.Range("B" & I2)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) Then
            
                .Range("E" & i).Formula = "=IF((B" & i & "- FLOOR(B" & i & ", ""1:00""))>J11,((B" & i & "- FLOOR(B" & i & ", ""1:00""))-J11),"""")" 'Insert cell E(i) the diff' between Timestamp and Roundown (01:00:00) minus the delta
                                                                                                                                                    'if the diff' between Timestamp and Roundown is bigger then delta.
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- Hour(i) = (Hour(i+1)-1) && Hour(i) = Hour(i-1) ----//////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- example: 21:00:00 = 21:00:00 And 21:00:00 = 21:00:00 ----///////////////////////////////////////////////
'//////////////////////////////////////////////////////---- extra: Hour(i) = (Hour(i+1)-1) - difference no more then one hour. ----////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ElseIf Hour(.Range("B" & i)) = Hour(.Range("B" & I1)) - 1 And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
                
                .Range("E" & i).Formula = "=IF(((Ceiling(B" & i & ",""1:00"")-B" & i & _
                    ")+(B" & i & "-B" & I2 & "))>J11,((Ceiling(B" & i & ", ""1:00"")-B" & i & ")+(B" & i & " - B" & I2 & ")-J11),"""")" 'Insert cell E(i) [ the diff' between Roundup (02:00:00) and Timestamp(i) plus the diff' between Timestamp(i) and Timestamp(i-1) minus the delta ]
                                                                                                                                        'if [ the diff' between Roundup and Timestamp(i) plus the diff of Timestamp(i) and Timestamp(i-1)] is bigger then delta.
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- (Timestamp(i) - Timestamp(i-1)) > Delta && Hour(i+1) = Hour(i) && Hour(i) = Hour(i-1)----//////////////////////////////
'//////////////////////////////////////////////////////---- example: (27/01/2021 07:03:00)-(27/01/2021 07:02:00) > 00:00:15 And 07:00:00 = 07:00:00 And 07:00:00 = 07:00:00----///
'//////////////////////////////////////////////////////---- extra: 07:00:00 the day after!. ----/////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

             ElseIf .Range("B" & i) - .Range("B" & I2) > .Range("J11") And Hour(.Range("B" & I1)) = Hour(.Range("B" & i)) And Hour(.Range("B" & i)) = Hour(.Range("B" & I2)) Then
            
                .Range("E" & i).Value = (.Range("B" & i).Value - .Range("B" & I2).Value) - .Range("J11").Value 'Insert cell E(i) [the diff' between Timestamp(i) minus Timestamp(i-1) minus the delta ].
                
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////---- No delta needed ----/////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Else
            
                .Range("E" & i).Value = ""
            
            End If
                
        End With
        
    Next i
    
End Sub


