Attribute VB_Name = "scroll_bar"
Private LastRow As Integer

Sub scroll_bar_Item(ws1, len1)

    LastRow = len1
    
    With ws1
    
        .Range("P2:P24").Formula = "=INDEX(B2:$B$" & (LastRow + 1) & ",$H$5)"
        .Range("Q2:Q24").Formula = "=INDEX(C2:$C$" & (LastRow + 1) & ",$H$5)"
        .Range("V2:V24").Formula = "=INDEX(E2:$E$" & (LastRow + 1) & ",$H$5)"
        .Range("D2").Value = ""
        .Range("E2").Value = ""
        
    End With
    
End Sub

Sub scroll_bar_Box(ws2, len2)
    
    LastRow = len2
    
    With ws2
    
        .Range("P2:P24").Formula = "=INDEX(B2:$B$" & (LastRow + 1) & ",$H$5)"
        .Range("Q2:Q24").Formula = "=INDEX(C2:$C$" & (LastRow + 1) & ",$H$5)"
        .Range("V2:V24").Formula = "=INDEX(E2:$E$" & (LastRow + 1) & ",$H$5)"
        .Range("D2").Value = ""
        .Range("E2").Value = ""
        
    End With
    
End Sub

Sub scroll_bar_Pallet(ws3, len3)
    
    LastRow = len3
    
    With ws3
    
        .Range("P2:P11").Formula = "=INDEX(B2:$B$" & (LastRow + 1) & ",$H$5)"
        .Range("Q2:Q11").Formula = "=INDEX(C2:$C$" & (LastRow + 1) & ",$H$5)"
        .Range("V2:V11").Formula = "=INDEX(E2:$E$" & (LastRow + 1) & ",$H$5)"
        .Range("D2").Value = ""
        .Range("E2").Value = ""
        
    End With
    
End Sub
