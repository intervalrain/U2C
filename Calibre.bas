
    Dim tempStr As String
    Dim AryStr
    Dim i As Integer
    
    Select Case Operator
        Case "N+", "P+", "C+"
            If InStr(1, Str, Operator) Then
                tempStr = Trim(Replace(Str, Operator, Replace(Operator, "+", "PLUS")))
                Calibre = tempStr
            Else
                Calibre = Str
            End If
    Case "N-", "P-", "C-"
            If InStr(1, Str, Operator) Then
                tempStr = Trim(Replace(Str, Operator, Replace(Operator, "-", "MINUS")))
                Calibre = tempStr
            Else
                Calibre = Str
            End If
        Case "AND"
            If InStr(1, Str, " AND ") Then
                Calibre = Trim(Replace(Str, " AND ", " and "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "NOT"
            If InStr(1, Str, " NOT ") Then
                Calibre = Trim(Replace(Str, " NOT ", " not "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "OR"
            If InStr(1, Str, " OR ") Then
                Calibre = Trim(Replace(Str, " OR ", " or "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "XOR"
            If InStr(1, Str, " XOR ") Then
                Calibre = Trim(Replace(Str, " XOR ", " xor "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "NOT_INTERACT"
            If InStr(1, Str, " NOT_INTERACT ") Then
                Calibre = Trim(Replace(Str, " NOT_INTERACT ", " not interact "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "INTERACT"
            If InStr(1, Str, " INTERACT ") Then
                Calibre = Trim(Replace(Str, " INTERACT ", " interact "))
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "SIZING"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " ")
                Str = Replace(Str, "um/side", " ")
                Str = Replace(Str, " +", " size by ")
                Str = Replace(Str, " -", " size by -")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "GROW"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Str = Replace(Str, " um", "")
                Str = Replace(Str, "RIGHT_BY", "right by")
                Str = Replace(Str, "LEFT_BY", "left by")
                Str = Replace(Str, "TOP_BY", "top by")
                Str = Replace(Str, "BOTTOM_BY", "bottom by")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "SHRINK"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Str = Replace(Str, " um", "")
                Str = Replace(Str, "RIGHT_BY", "right by")
                Str = Replace(Str, "LEFT_BY", "left by")
                Str = Replace(Str, "TOP_BY", "top by")
                Str = Replace(Str, "BOTTOM_BY", "bottom by")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "REVERSE"
            If InStr(1, Str, " " & Operator) Then
                If IncludedOrNot Then
                    Calibre = Replace(Str, Operator, Window & " not ")
                    IncludedOrNot = True
                ElseIf IncludedOrNot = False Then
                    Calibre = getCOL(Str, Operator, 2)
                    Calibre = Replace(Calibre, "(", "")
                    Calibre = Replace(Calibre, ")", "")
                    Calibre = Window & " not " & Calibre
                    Calibre = Trim(getCOL(Str, Operator, 1) & " " & Calibre)
                    IncludedOrNot = True
                End If
            Else
                Calibre = Str
            End If
        Case "AREA"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Str = Replace(Str, "um^2", " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "HOLES"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "RECTANGLE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Str = Replace(Str, " BY ", " by ")
                Str = Replace(Str, " um", "")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "NOT_OUTSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "OUTSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "NOT_INSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case "INSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " " & LCase(Operator) & " ")
                Calibre = Trim(Str)
                IncludedOrNot = True
            Else
                Calibre = Str
            End If
        Case Else
            Calibre = Str
    End Select
    
End Function
