
Public Function init(ByVal strToBeInit As String) As String

    Dim i As Integer

    If InStr(strToBeInit, "//") Then strToBeInit = Trim(getCOL(strToBeInit, "//", 1))
    
    strToBeInit = Replace(strToBeInit, "==", "eq")
    strToBeInit = Replace(strToBeInit, "!=", "ne")
    strToBeInit = Replace(strToBeInit, ">=", "ge")
    strToBeInit = Replace(strToBeInit, "<=", "le")
    strToBeInit = Replace(strToBeInit, "=", " = ")
    
    For i = 1 To 7
        strToBeInit = Replace(strToBeInit, "  ", " ")
        strToBeInit = Replace(strToBeInit, "( ", "(")
        strToBeInit = Replace(strToBeInit, " )", ")")
    Next i
    strToBeInit = Trim(getCOL(strToBeInit, "=", 1)) & " = " & Trim(getCOL(strToBeInit, "=", 2))
    
    
    strToBeInit = Replace(strToBeInit, "eq", "==")
    strToBeInit = Replace(strToBeInit, "ne", "!=")
    strToBeInit = Replace(strToBeInit, "ge", ">=")
    strToBeInit = Replace(strToBeInit, "le", "<=")
    init = strToBeInit
    
End Function

Public Function TrimLn(ByVal Str As String)

    Dim tempStr As String
    
    tempStr = Str
    tempStr = Replace(Str, Chr(10), "")
    tempStr = Replace(Str, Chr(13), "")
    tempStr = Trim(tempStr)
    TrimLn = tempStr

End Function

Public Function addTail(ByVal Str As String, ByVal keyword As String, ByVal tail As Integer)

    Dim tempAry() As String
    Dim i As Integer
    Dim tempStr As String

    
    If InStr(1, Str, keyword) Then
        tempAry = Split(Str, " ")
        For i = 0 To UBound(tempAry)
            If InStr(1, tempAry(i), keyword) Then
                If InStr(1, tempAry(i), "REVERSE") > 0 Then
                    tempStr = Replace(tempAry(i), "REVERSE(", "")
                    tempStr = Replace(tempStr, ")", "")
                    tempAry(i) = Replace(tempAry(i), tempStr, tempStr & "_" & tail)
                Else
                    tempStr = Replace(tempAry(i), "(", "")
                    tempStr = Replace(tempStr, ")", "")
                    tempAry(i) = Replace(tempAry(i), tempStr, tempStr & "_" & tail)
                End If
            End If
        Next
    Else
        addTail = Str
        Exit Function
    End If
    
    addTail = Join(tempAry, " ")
    
End Function


Public Function transRecordWord(mStr As String, ByRef strMap)
    Dim v
    Dim extraCheckStr
    For Each v In strMap.keys
        If InStr(mStr, v) Then
            mStr = Replace(mStr, v, strMap(v))
            mStr = trimParentheses(mStr)
        End If
    Next v
    
    extraCheckStr = Split(getCOL(mStr, " = ", 2))
    If UBound(extraCheckStr) > 1 Then
        If UBound(extraCheckStr) = 2 And extraCheckStr(0) = extraCheckStr(2) Then
            If Not strMap.exists(getCOL(mStr, " = ", 1)) Then
                strMap.Add getCOL(mStr, " = ", 1), "(" & extraCheckStr(0) & ")"
            End If
            mStr = ""
        End If
    End If
    
    transRecordWord = mStr

End Function

Public Function trimParentheses(mStr As String)

    Dim S As Integer, E As Integer, i As Integer
    Dim tmpStr As String
    Dim curr As String
    Dim flag As Boolean

    tmpStr = mStr
    For i = 1 To Len(mStr)
        curr = Mid(mStr, i, 1)
        If curr = "(" Then S = i
        If curr = " " Then S = 0
        If curr = ")" Then E = i
        
        If S > 0 And E > 0 And E > S Then
            Dim s1 As String
            Dim s2 As String
            Dim s3 As String
            s1 = Left(mStr, S - 1)
            s3 = Right(mStr, Len(mStr) - E)
            s2 = Mid(mStr, S + 1, Len(mStr) - Len(s3) - Len(s1) - 2)
            tmpStr = s1 & s2 & s3
            flag = True
            Exit For
        End If
    Next i

    If flag Then trimParentheses = trimParentheses(tmpStr)
    trimParentheses = tmpStr

End Function

Public Function TrimLn(ByVal Str As String)

    Dim tempStr As String
    
    tempStr = Str
    tempStr = Replace(Str, Chr(10), "")
    tempStr = Replace(Str, Chr(13), "")
    tempStr = Trim(tempStr)
    TrimLn = tempStr

End Function
