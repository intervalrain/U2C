Option Explicit

Public Function getCOL(ByVal mString As String, ByVal SplitChar As String, ByVal mNumber As Long)
   Dim tempA
   tempA = Split(mString, SplitChar)
   If mNumber - 1 > UBound(tempA) Or mNumber < 1 Then
      getCOL = ""
   Else
      getCOL = Trim(tempA(mNumber - 1))
   End If
End Function

Public Function AddSheet(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Integer
   Application.DisplayAlerts = False
   For i = 1 To Worksheets.Count
      If UCase(Worksheets(i).Name) = UCase(sheetName) Then Worksheets(i).Delete: Exit For
   Next
   Set nowSheet = Worksheets.Add(, Worksheets(Worksheets.Count))
   nowSheet.Name = sheetName
   Set AddSheet = nowSheet
   Set nowSheet = Nothing
   Application.DisplayAlerts = True
End Function

Public Function DelSheet(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Integer
   Application.DisplayAlerts = False
   For i = 1 To Worksheets.Count
      If Worksheets(i).Name = sheetName Then Worksheets(i).Delete: Exit For
   Next
   Application.DisplayAlerts = True
End Function


Public Function IsExistSheet(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Integer
   
   For i = 1 To Worksheets.Count
      If UCase(Worksheets(i).Name) = UCase(sheetName) Then IsExistSheet = True: Exit Function
   Next
   IsExistSheet = False
End Function

Public Function ReadTextFile(mFilename As String)
    Dim FSO As New Scripting.FileSystemObject
    Dim F As TextStream
    
    Set F = FSO.OpenTextFile(mFilename, 1)   'ForReading
    ReadTextFile = F.ReadAll
    F.Close
    Set FSO = Nothing
    Set F = Nothing
End Function

Public Function GetFileHeader(mFilename As String, mLine)
    Dim FileID As Long
    Dim temp As String, tempStr
    Dim i As Long
    
    FileID = FreeFile
    Open mFilename For Input As #FileID
         For i = 1 To mLine
             If Not EOF(FileID) Then Line Input #FileID, temp: tempStr = tempStr & temp & vbCrLf
         Next i
    Close #FileID
    GetFileHeader = tempStr
End Function


Public Function FileDialog(FileType As String)
    Dim tempAry
    Dim i As Long
    Dim FilterStr As String
    Dim tmpStr As String
    
    tempAry = Split(FileType, ",")
    For i = 0 To UBound(tempAry)
        FilterStr = FilterStr & "," & tempAry(i) & " File(*." & tempAry(i) & "),*." & tempAry(i)
    Next i
    FilterStr = FilterStr & "," & "All File(*.*),*.*"
    FilterStr = Mid(FilterStr, 2)
   
    tmpStr = Application.GetOpenFilename(FilterStr, 1, "Open File", "Open", False)
    If UCase(tmpStr) = "FALSE" Then
        FileDialog = ""
    Else
        FileDialog = tmpStr
    End If
End Function


Public Function N2L(ByVal mNum As Integer) As String
    Select Case mNum
        Case 0
            N2L = "O"
        Case 1
            N2L = "I"
        Case 2
            N2L = "R"
        Case 3
            N2L = "E"
        Case 4
            N2L = "A"
        Case 5
            N2L = "S"
        Case 6
            N2L = "G"
        Case 7
            N2L = "T"
        Case 8
            N2L = "B"
        Case 9
            N2L = "Q"
    End Select
End Function

Public Function IsInArray(ByVal stringToBeFound As String, ByVal Arr As Variant) As String
    If InStr("^" & Join(Arr, "^") & "^", stringToBeFound) = 0 Then
        IsInArray = ""
    Else
        IsInArray = Replace(stringToBeFound, "^", "")
    End If
    
End Function
