
Public Function DRCS(ByVal Str As String, ByVal Operator As String, ByVal Window As String, ByRef IncludedOrNot As Boolean)

    Dim tempStr As String
    Dim AryStr
    Dim i As Integer
    Dim L1 As String, L2 As String, L3 As String
    Dim Target As String
    Dim Result As String
    
    Select Case Operator
        Case "N+", "P+", "C+"
            If InStr(1, Str, Operator) Then
                tempStr = Trim(Replace(Str, Operator, Replace(Operator, "+", "PLUS")))
                DRCS = tempStr
            Else
                DRCS = Str
            End If
        Case "N-", "P-", "C-"
            If InStr(1, Str, Operator) Then
                tempStr = Trim(Replace(Str, Operator, Replace(Operator, "-", "MINUS")))
                DRCS = tempStr
            Else
                DRCS = Str
            End If
        Case "AND"
            If InStr(1, Str, " AND ") Then
                DRCS = Trim(Replace(Str, " AND ", " & "))
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "NOT"
            If InStr(1, Str, " NOT ") Then
                DRCS = Trim(Replace(Str, " NOT ", " - "))
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "OR"
            If InStr(1, Str, " OR ") Then
                DRCS = Trim(Replace(Str, " OR ", " | "))
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "XOR"
            If InStr(1, Str, " XOR ") Then
                DRCS = Trim(Replace(Str, " XOR ", " ^ "))
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "NOT_INTERACT"
            If InStr(1, Str, " NOT_INTERACT ") Then
                AryStr = Split(getCOL(Str, " = ", 2), " NOT_INTERACT ")
                For i = 1 To UBound(AryStr)
                    AryStr(i) = AryStr(i) & ")"
                Next i
                DRCS = getCOL(Str, " = ", 1) & " = " & Join(AryStr, ".not_interacting(")
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "INTERACT"
            If InStr(1, Str, " INTERACT ") Then
                AryStr = Split(getCOL(Str, " = ", 2), " INTERACT ")
                For i = 1 To UBound(AryStr)
                    AryStr(i) = AryStr(i) & ")"
                Next i
                DRCS = getCOL(Str, " = ", 1) & " = " & Join(AryStr, ".interacting(")
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "SIZING"
            If InStr(1, Str, " SIZING ") Then
                AryStr = Split(getCOL(Str, " um/side", 1), " ")
                AryStr(UBound(AryStr)) = Replace(AryStr(UBound(AryStr)), "+", "")
                tempStr = getCOL(getCOL(Str, "SIZING ", 2), " ", 1)
                tempStr = tempStr & ".sized(" & AryStr(UBound(AryStr)) & ")"
                tempStr = getCOL(Str, " = ", 1) & " = " & tempStr
                DRCS = Trim(tempStr)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "GROW"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " ")
                If InStr(1, Str, "RIGHT_BY") Then
                    Str = Trim(getCOL(Str, "RIGHT_BY", 1)) & ".sized(" & Trim(getCOL(getCOL(Str, "RIGHT_BY", 2), "um", 1)) & ",0)"
                ElseIf InStr(1, Str, "LEFT_BY") Then
                    Str = Trim(getCOL(Str, "LEFT_BY", 1)) & ".sized(" & Trim(getCOL(getCOL(Str, "LEFT_BY", 2), "um", 1)) & ",0)"
                ElseIf InStr(1, Str, "TOP_BY") Then
                    Str = Trim(getCOL(Str, "TOP_BY", 1)) & ".sized(0," & Trim(getCOL(getCOL(Str, "TOP_BY", 2), "um", 1)) & ")"
                ElseIf InStr(1, Str, "BOTTOM_BY") Then
                    Str = Trim(getCOL(Str, "BOTTOM_BY", 1)) & ".sized(0," & Trim(getCOL(getCOL(Str, "BOTTOM_BY", 2), "um", 1)) & ")"
                End If
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "SHRINK"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " ")
                If InStr(1, Str, "RIGHT_BY") Then
                    Str = Trim(getCOL(Str, "RIGHT_BY", 1)) & ".sized(" & -1 * Trim(getCOL(getCOL(Str, "RIGHT_BY", 2), "um", 1)) & ",0)"
                ElseIf InStr(1, Str, "LEFT_BY") Then
                    Str = Trim(getCOL(Str, "LEFT_BY", 1)) & ".sized(" & -1 * Trim(getCOL(getCOL(Str, "LEFT_BY", 2), "um", 1)) & ",0)"
                ElseIf InStr(1, Str, "TOP_BY") Then
                    Str = Trim(getCOL(Str, "TOP_BY", 1)) & ".sized(0," & -1 * Trim(getCOL(getCOL(Str, "TOP_BY", 2), "um", 1)) & ")"
                ElseIf InStr(1, Str, "BOTTOM_BY") Then
                    Str = Trim(getCOL(Str, "BOTTOM_BY", 1)) & ".sized(0," & -1 * Trim(getCOL(getCOL(Str, "BOTTOM_BY", 2), "um", 1)) & ")"
                End If
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "REVERSE"
            If InStr(1, Str, " REVERSE") Then
                If IncludedOrNot Then
                    DRCS = Replace(Str, Operator, Window & " - ")
                    IncludedOrNot = True
                ElseIf IncludedOrNot = False Then
                    DRCS = getCOL(Str, Operator, 2)
                    DRCS = Replace(DRCS, "(", "")
                    DRCS = Replace(DRCS, ")", "")
                    DRCS = Window & " - " & DRCS
                    DRCS = Trim(getCOL(Str, Operator, 1) & " " & DRCS)
                    IncludedOrNot = True
                End If
            Else
                DRCS = Str
            End If
        Case "AREA"
            Dim constraint As String
            Dim UBd As String
            Dim LBd As String
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " ")
                constraint = getCOL(getCOL(Str, " = ", 2), " ", 2)
                If InStr(constraint, "<") Then
                    UBd = "<" & getCOL(constraint, "<", 2)
                    LBd = getCOL(constraint, "<", 1)
                Else
                    LBd = constraint
                End If
                If Not UBd = "" And Not LBd = "" Then
                    Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select{|polygon| polygon.area " & LBd & "} | " & getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select{|polygon| polygon.area " & UBd & "}"
                ElseIf Not UBd = "" Then
                    Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select{|polygon| polygon.area " & UBd & "}"
                ElseIf Not LBd = "" Then
                    Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select{|polygon| polygon.area " & LBd & "}"
                End If
                
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "HOLES"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = getCOL(Str, " = ", 1) & " = " & getCOL(Str, " HOLES ", 2) & ".holes"
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "RECTANGLE"
            Dim VBd As String
            Dim HBd As String
            Dim Head As String
            If InStr(1, Str, " " & Operator & " ") Then
                Str = Replace(Str, " " & Operator & " ", " ")
                Str = Replace(Str, " um", "")
                Str = Replace(Str, " BY ", " ")
                constraint = Mid(getCOL(Str, " = ", 2), InStr(getCOL(Str, " = ", 2), " ") + 1)
                VBd = getCOL(constraint, " ", 1) & getCOL(constraint, " ", 2)
                HBd = getCOL(constraint, " ", 3) & getCOL(constraint, " ", 4)
                Head = getCOL(Str, " = ", 1) & " = "
                If VBd = HBd Then
                    Str = getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select {|polygon| polygon.bbox.width &tmpStr }"
                    Str = Replace(Str, "&tmpStr", VBd) & " & " & Replace(Replace(Str, "bbox.width", "bbox.height"), "&tmpStr", HBd)
                    DRCS = Trim(Head & Str)
                Else
                    Str = getCOL(getCOL(Str, " = ", 2), " ", 1) & ".select {|polygon| polygon.bbox.width &tmpStr }"
                    Str = Replace(Str, "&tmpStr", "&tmpStr1") & " & " & Replace(Replace(Str, "bbox.width", "bbox.height"), "&tmpStr", "&tmpStr2")
                    Str = "(" & Replace(Replace(Str, "&tmpStr1", VBd), "&tmpStr2", HBd) & ") | (" & Replace(Replace(Str, "&tmpStr1", HBd), "&tmpStr2", VBd) & ")"
                    DRCS = Trim(Head & Str)
                End If
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "NOT_OUTSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 1) & "." & LCase(Operator) & "(" & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 2) & ")"
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "OUTSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 1) & "." & LCase(Operator) & "(" & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 2) & ")"
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "NOT_INSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 1) & "." & LCase(Operator) & "(" & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 2) & ")"
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "INSIDE"
            If InStr(1, Str, " " & Operator & " ") Then
                Str = getCOL(Str, " = ", 1) & " = " & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 1) & "." & LCase(Operator) & "(" & getCOL(getCOL(Str, " = ", 2), " " & Operator & " ", 2) & ")"
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        
        'For DRC function
        Case "TOUCH"
            If InStr(1, Str, " " & Operator & " ") Then
                
                AryStr = Split(Str, " ")
                L3 = AryStr(0)
    
                If InStr(Str, Operator) And UBound(AryStr) > 3 Then
                    For i = 3 To UBound(AryStr) - 1
                        If AryStr(i) = Operator Then
                            L1 = Replace(AryStr(i - 1), "(", "")
                            L2 = Replace(AryStr(i + 1), ")", "")
                            Exit For
                        End If
                    Next i
                End If
            
                Target = L1 & " " & Operator & " " & L2
                Result = L1 & ".interacing(" & L2 & ") & " & L1 & ".outside(" & L2 & ")"
                Str = Replace(Str, Target, Result)
                
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
            
        Case "CUT"
            If InStr(1, Str, " " & Operator & " ") Then
                
                AryStr = Split(Str, " ")
                L3 = AryStr(0)
    
                If InStr(Str, Operator) And UBound(AryStr) > 3 Then
                    For i = 3 To UBound(AryStr) - 1
                        If AryStr(i) = Operator Then
                            L1 = Replace(AryStr(i - 1), "(", "")
                            L2 = Replace(AryStr(i + 1), ")", "")
                            Exit For
                        End If
                    Next i
                End If
            
                Target = L1 & " " & Operator & " " & L2
                Result = L1 & ".interacing(" & L2 & ") - " & L1 & ".inside(" & L2 & ") - " & L1 & ".outside(" & L2 & ")"
                Str = Replace(Str, Target, Result)
                
                DRCS = Trim(Str)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case "DONUT"
            If InStr(1, Str, Operator) Then
                
                AryStr = Split(Str, " ")
                
                If InStr(Str, Operator) And UBound(AryStr) > 2 Then
                    For i = 2 To UBound(AryStr) - 1
                        If Replace(AryStr(i), "(", "") = Operator Then
                            L1 = Replace(AryStr(i + 1), ")", "")
                            Exit For
                        End If
                    Next i
                End If
                
                Target = Operator & " " & L1
                Result = L1 & ".interacing(" & L1 & ".holes" & ")"
                
                
                DRCS = Replace(Str, Target, Result)
                IncludedOrNot = True
            Else
                DRCS = Str
            End If
        Case Else
            DRCS = Str
            
    End Select
    
End Function
