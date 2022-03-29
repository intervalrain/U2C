' ****************************************************************************************************
' *
' * @Macro      DRCS
' * @Version    1.13
' * @New Update 2021/12/14
' *
' * @Description
' *             To translate boolean operation code from MT Form format to KLayout DRC script format,
' *             Due to variation among MT Form format, Calibre and DRC codes(Ruby) in KLayout,
' *             It's not official procedure for try run boolean.
' *             It could ONLY be REFERENCE for boolean operation development duration.
' *             Please confirm post-boolean result on mask JDV system.
' *
' * @Author     Rain Hu
' * @Mail       rain_hu@umc.com
' * @Dept       TD3-SD-DT4
' * @Extension  37187
' *
' ****************************************************************************************************

Public flag As Boolean
Public Window As String
Public Keywords
Public subType

Option Explicit
' ****************************************************************************************************
' *
' * @Sub        AutoRun
' *
' * @Description
' *     To connect all subs automatically.
' *
' ****************************************************************************************************
Public Sub AutoRun()

    If Not IsExistSheet("CtrlTable") Then MsgBox ("Sheet ""CtrlTable"" is not found."): Exit Sub
    If Not IsExistSheet("LayerTable") Then MsgBox ("Sheet ""LayerTable"" is not found."): Exit Sub
    
    Call genKeywordList
    Call Initial
        If flag Then Exit Sub
    Call Run
        If flag Then Exit Sub
    Call Layer
        If flag Then Exit Sub
    Call Result

End Sub

' ****************************************************************************************************
' *
' * @Sub        AutoRun_Calibre
' *
' * @Description
' *     To connect all subs automatically.
' *
' ****************************************************************************************************
Public Sub AutoRun_Calibre()

    If Not IsExistSheet("CtrlTable") Then MsgBox ("Sheet ""CtrlTable"" is not found."): Exit Sub
    If Not IsExistSheet("LayerTable") Then MsgBox ("Sheet ""LayerTable"" is not found."): Exit Sub
    
    'Call genKeywordList
    Call Initial
        If flag Then Exit Sub
    Call Run("Calibre")
        If flag Then Exit Sub
    Call Layer("Calibre")
        If flag Then Exit Sub
    Call Result("Calibre")

End Sub

Private Sub genKeywordList()
' ****************************************************************************************************
' *
' * @Sub        genKeywordList
' *
' * @Description
' *             To avoid prefix of layer name being numeric.
' *             Numbers will be transfer into letters shown as below.
' *             0 -> O
' *             1 -> I
' *             2 -> R
' *             3 -> E
' *             4 -> A
' *             5 -> S
' *             6 -> G
' *             7 -> T
' *             8 -> B
' *             9 -> Q
' *
' ****************************************************************************************************
    
    Dim inSheet As Worksheet
    Dim i As Long, j As Long
    ReDim Keywords(0)
    Set inSheet = Worksheets("LayerTable")
    
    For i = 2 To inSheet.UsedRange.Rows.Count
        If IsNumeric(Left(inSheet.Cells(i, 1), 1)) Then
            Keywords(j) = inSheet.Cells(i, 1)
            j = j + 1
            ReDim Preserve Keywords(j)
        End If
    Next i
    ReDim Preserve Keywords(j - 1)
    Set inSheet = Nothing
End Sub

Private Sub Initial()
' ****************************************************************************************************
' *
' * @Sub        Initial
' *
' * @Description
' *             To split sequences in merged rows into separate rows.
' *             New sheet named "MES" will be produced in this step.
' *             The sheet "MES" could be used for manual function "MergeRows".
' *
' ****************************************************************************************************
    
    Dim inSheet As Worksheet        'Sheet CtrlTable
    Dim outSheet As Worksheet       'Sheet MES
    Dim tempAry() As String
    Dim i As Long, j As Long        '°Ï°ìÅÜ¼Æ
    Dim LineCount As Integer
    Dim Highlight As Integer        'Flagªº¦æ¼Æ
    
    '====================§PÂ_¦³¨S¦³CtrlTable====================
    If Not IsExistSheet("CtrlTable") Then MsgBox ("Sheet ""CtrlTable"" is not found."): Exit Sub
    
    Set inSheet = Worksheets("CtrlTable")
    Set outSheet = AddSheet("MES")
    LineCount = 1
    flag = False
    '====================±NMTFormªº®æ¦¡Âà´«¦¨³v¦æ====================
    For i = 1 To inSheet.UsedRange.Rows.Count
        If inSheet.Cells(i, 1).Value = "" Or UCase(inSheet.Cells(i, 1).Value) = "END" Then Exit For     '¹J¨ìªÅ¥Õ©ÎEND§Y°±¤î
        tempAry = Split(inSheet.Cells(i, 2).Value, Chr(10))                                             '¥H´«¦æChr(10)§@¬°Delimiter±N¤º®e¦s¤J¯x°}
        For j = 0 To UBound(tempAry)
            tempAry(j) = TrimLn(tempAry(j))                                                             '°£¥h¦h¾lªºªÅ¥Õ¦r¤¸
            outSheet.Cells(j + LineCount, 2).Value = tempAry(j)                                         '¨Ã¦bSheet MES¦L¥X
            If UBound(Split(tempAry(j), "(")) <> UBound(Split(tempAry(j), ")")) Then
                outSheet.Cells(j + LineCount, 2).Interior.ColorIndex = 3
                flag = True
                Highlight = j + LineCount
            End If
        Next
            outSheet.Cells(LineCount, 1).Value = getCOL(inSheet.Cells(i, 1).Value, Chr(10), 1)      '°£¥h´«¦æ¥H«áªº¦r¦ê, ¥u¨úLayerName
        LineCount = outSheet.UsedRange.Rows.Count + 2
    Next
    
    '====================®ø°£­«½ÆªÅ¥Õ¦æ====================
    For i = outSheet.UsedRange.Rows.Count To 1 Step -1
        If outSheet.Cells(i, 2).Value = "" Then
            If outSheet.Cells(i + 1, 2).Value = "$VARIABLE_START" Or outSheet.Cells(i - 1, 2).Value = "$VARIABLE_END" Then
            Else
                outSheet.Rows(i).Delete
            End If
            If i < Highlight Then Highlight = Highlight - 1
        End If
    Next
    
    If flag Then MsgBox ("¦³Booelan¦¡ªº¥ª¥k¬A¸¹¤£¹ïºÙ"): Cells(Highlight, 2).Activate
    
    Set inSheet = Nothing
    Set outSheet = Nothing

End Sub

Private Sub Run(Optional Mode As String = "DRCS")
' ****************************************************************************************************
' *
' * @Sub        Run
' *
' * @Description
' *             To transform codes from MT form format into KLayout DRC Script format(ruby).
' *
' * @Function1  Add suffix in layer name to prevent duplicated naming with different definition.
' * @Function2  Highlight error of boolean codes for users to debug.
' *             (1) Duplicated layer name
' *             (2) Prefix with numbers
' *             (3) Asymmetic parentheses
' *
' ****************************************************************************************************

    Dim inSheet As Worksheet        'Sheet MES
    Dim outSheet As Worksheet       'Sheet DRCS
    Dim tSheet As Worksheet         'Sheet TEMP
    
    Dim mStr As String              '¦r¦êÅÜ¼Æ(Boolean¦¡)
    Dim tStr As String              '¦r¦êÅÜ¼Æ(Layer Name)
    Dim vStr As Variant             'Boolean¦¡¤¤ªºÅÜ¼Æ©w¸qÃöÁä¦r¸s
    Dim v As Variant                'Boolean¦¡¤¤ªºÅÜ¼Æ©w¸qÃöÁä¦r
    
    Dim i As Long, j As Long        '°Ï°ìÅÜ¼Æ
    
    Dim tail As Integer             '¥Î¦b¦r§ÀÁ×§KTEMP¦r¦ê­«½Æ(¼Æ¦r)

    Dim AryOperator As Variant      'Calibre Operator
    
    Dim ActOrNot As Boolean         '¦³¨S¦³³QCalibre Function³B²z¹L

    Dim tempAry                     'ÀË¬dprefix¨Ï¥Îªº¯x°}
    Dim tempStr                     'ÀË¬dprefix¨Ï¥Îªº«O¯d¦r
    Dim extraCheckStr               'ÀË¬d©w¸q¦¡¬O§_¦h¾l
    
    Dim AryCheck As Variant         'Àx¦sÀË¬d­«½Æªº¯x°}
    Dim MsgAlarm As String          '´£¥Ü¦r¦ê
    Dim PtToLine As Integer         '²¾°Ê¨ìªº¥Ø¼Ð¦æ¼Æ
    
    Dim strMap                      '¬ö¿ý¨ú¥N¦r
    Set strMap = CreateObject("Scripting.Dictionary")
        
    '====================Initilization====================
    
    Set inSheet = Worksheets("MES")
    Set outSheet = AddSheet(Mode)

    '====================¿é¤JWindow Layer Name====================
    Window = InputBox("Window=", "Reverse Setting", "SUBSTRATE")
    If Window = "" Then flag = True: Exit Sub                                                           '¨ú®ø«hÂ÷¶}µ{¦¡
    '====================¿é¤JWindow Layer Name====================
    subType = InputBox("Add a constant int(0~255) to output layers type", "Sub-Datatype setting", 100)
    If subType = "" Then subType = 0
    If Not IsNumeric(subType) Then flag = True: MsgBox ("Please input a int between 0 and 255"): Exit Sub
    If subType < 0 Or subType > 255 Then flag = True: MsgBox ("Please input a int between 0 and 255"): Exit Sub
    '====================DRCscript Transform====================
    For i = 1 To inSheet.UsedRange.Rows.Count

        mStr = inSheet.Cells(i, 2).Value
        '==========Initialization==========
        If inSheet.Cells(i, 1).Value <> "" Then
            tStr = inSheet.Cells(i, 1).Value
            tail = tail + 1                                                                             '¦r§ÀÀH¤£¦PLayerªº°µAuto increment
        End If
        mStr = init(mStr)
        '==========¬Ù²¤Boolean¦¡¤¤ªºÅÜ¼Æ©w¸q==========
        vStr = Array("$VARIABLE_START", _
                     "$VARIABLE_END", _
                     "$PARAMETER_START", _
                     "$PARAMETER_END", _
                     "a customized_data provided at tape out by customer", _
                     "an operation_data outputted together with OPC data")
           
        For Each v In vStr
            If InStr(1, mStr, v) Then mStr = ""
        Next
        '==========³B²zDUMMY/OPC¦r¦ê==========                                                          '2021/9/13
        If mStr Like "*DUMMY(*)" Then
            If Not strMap.exists(getCOL(mStr, " = ", 1)) Then
                strMap.Add getCOL(mStr, " = ", 1), "(" & getCOL(mStr, "DUMMY(", 2)
            End If
            mStr = ""
        End If

        If mStr Like "*OPC(*)" Then
            If Not strMap.exists(getCOL(mStr, " = ", 1)) Then
                strMap.Add getCOL(mStr, " = ", 1), "(" & getCOL(mStr, "OPC(", 2)
            End If
            mStr = ""
        End If
        
        mStr = transRecordWord(mStr, strMap)
        
        '==========¥[¦r§À==========
        mStr = addTail(mStr, "TEMP", tail)
        '==========Prefix Âà´«==========  '211214 §ó·s
        If Mode = "DRCS" Then
            tempAry = Split(mStr, " ")
            For j = 0 To UBound(tempAry)
                tempStr = IsInArray("^" & Replace(Replace(tempAry(j), "(", ""), ")", "") & "^", Keywords)
                If Not tempStr = "" Then
                    tempAry(j) = Replace(tempAry(j), tempStr, N2L(Left(tempStr, 1)) & Mid(tempStr, 2))
                End If
            Next j
            mStr = Join(tempAry, " ")
        End If
        '==========DRCscript Âà´«==========
        AryOperator = Array("AND", _
                            "NOT", _
                            "OR", _
                            "XOR", _
                            "INTERACT", _
                            "NOT_INTERACT", _
                            "N+", _
                            "P+", _
                            "C+", _
                            "N-", _
                            "P-", _
                            "C-", _
                            "SIZING", _
                            "GROW", _
                            "SHRINK", _
                            "REVERSE", _
                            "AREA", _
                            "HOLES", _
                            "RECTANGLE", _
                            "OUTSIDE", _
                            "NOT_OUTSIDE", _
                            "INSIDE", _
                            "NOT_INSIDE")
        If Mode = "DRCS" Then
            For Each v In AryOperator
                mStr = DRCS(mStr, v, Window, ActOrNot)                                                   '¹ïAryOperator¤¤§t¦³ªº¹Bºâ¤l°µÂà´«(©I¥sSubFunction¤¤ªºDRCS)
            Next v
        ElseIf Mode = "Calibre" Then
            For Each v In AryOperator
                mStr = Calibre(mStr, v, Window, ActOrNot)                                                '¹ïAryOperator¤¤§t¦³ªº¹Bºâ¤l°µÂà´«(©I¥sSubFunction¤¤ªºCalibre)
            Next v
        End If
        '==========³B²zµL°Ê§@ªºBoolean¦¡==========
        If ActOrNot = False Then mStr = ""                                                              '­Y¨S¦³¸g¹LÂà´««h§R°£
        If Trim(getCOL(mStr, "=", 1)) = Trim(getCOL(mStr, "=", 2)) Then mStr = ""                       '­Yµ¥¸¹¨âÃä¬Û¦P«h§R°£
        '==========Print==========
        outSheet.Cells(i, 1).Value = mStr                                                               '¦C¦L
        ActOrNot = False                                                                                '¦L§¹«á­«¸mª¬ºA
    Next i
    '====================§R°£¦h¾lªÅ¥Õ¦C====================                                             '2021/9/13 Update
    For i = outSheet.UsedRange.Rows.Count To 2 Step -1
        
        If outSheet.Cells(i, 1) = "" And inSheet.Cells(i + 1, 1) = "" Then
            outSheet.Rows(i).Delete
        End If
                
    Next i
    '====================ÀË¬d¦³µL­«½Æ¦¡¤l====================
    AddSheet ("temp")
    
    Set tSheet = Worksheets("Temp")
    outSheet.Columns(1).Copy tSheet.Columns(2)                                                          '±NSheet Calibre½Æ»s¨ìSHeet TEMP
    
    For i = 1 To tSheet.UsedRange.Rows.Count + 1
        tSheet.Cells(i, 1).Value = CStr(i)                                                              '±N²Ä¤@Äæ¶ñ¤W¦C¼Æ
    Next
    '==========±N²Ä2Äæªº¤å¦r°µ­°§Ç±Æ¦C==========
    tSheet.Range("A:B").Sort Key1:=Range("B1"), Order1:=xlAscending, _
                             Key2:=Range("A1"), Order2:=xlAscending, _
                             Header:=xlGuess, Ordercustom:=1, MatchCase:=False, _
                             Orientation:=xlTopToBottom
                              
    '==========±N¤å¦r¥H"µ¥¸¹"°µ¸ê®Æ­åªR«á¤À¸Ñ==========
    tSheet.Columns(2).TextToColumns Destination:=Range("B1"), _
                                       DATATYPE:=xlDelimited, _
                                  TextQualifier:=xlDoubleQuote, _
                           ConsecutiveDelimiter:=False, _
                                            Tab:=False, _
                                      Semicolon:=False, _
                                          Comma:=False, _
                                          Space:=False, _
                                          Other:=True, _
                                      OtherChar:="=", _
                                      FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                           TrailingMinusNumbers:=True
   '==========ÀË¬d¦³µL­«½Æ==========
    flag = False
    For i = tSheet.UsedRange.Rows.Count To 2 Step -1
        If Trim(tSheet.Cells(i, 2)) <> "" And Trim(tSheet.Cells(i, 2)) = Trim(tSheet.Cells(i - 1, 2)) Then
            If Not Trim(tSheet.Cells(i, 3)) = Trim(tSheet.Cells(i - 1, 3)) Then
                outSheet.Rows(tSheet.Cells(i - 1, 1).Value).Interior.ColorIndex = 3                                 '±N­«½Æªº¦¡¤l¥Î¬õ¦âhighlight
                outSheet.Rows(tSheet.Cells(i, 1).Value).Interior.ColorIndex = 3
                tSheet.Cells(i, 4).Value = tSheet.Cells(i, 1).Value                                                 '¦bsheet TEMP¤¤°µ¼Ð°O
                tSheet.Cells(i - 1, 4).Value = tSheet.Cells(i - 1, 1).Value                                         '¦bsheet TEMP¤¤°µ¼Ð°O
                flag = True                                                                                         '°µerror¼Ð°O ¨Ïµ{¦¡¼È°±
            Else
                outSheet.Rows(tSheet.Cells(i, 1).Value).Delete                                                      '­Y©w¸q¬Û¦P«h§R°£
                tSheet.Rows(i).Delete
            End If
        End If
    Next
   
   '==========±N¤å¦r«ì´_­ì¥»ªº±Æ¦C==========
    tSheet.Range("A:D").Sort Key1:=Range("A1"), Order1:=xlAscending, _
                             Header:=xlGuess, Ordercustom:=1, MatchCase:=False, _
                             Orientation:=xlTopToBottom
   '==========¨¾§b(­«½Æ)´£¥Ü==========
    If flag Then                                                                                                '­Yerror¼Ð°O¬°true«hÅã¥Ü´£¿ô
        MsgAlarm = "µ{§Ç¤¤Â_, ¦]¥H¤Uªº¦C¼Æªº©R¦W¦³­«½Æ¨Ï¥Î, ½Ð¤â°Ê°£¿ù" & Chr(13)
        ReDim AryCheck(tSheet.UsedRange.Rows.Count + 1) As String
        For i = 1 To tSheet.UsedRange.Rows.Count
            AryCheck(i) = tSheet.Cells(i, 4).Value                                                              '±N¼Ð°OÄæ¦s¤JAryCheck
        Next
        For Each v In AryCheck
            If Not v = "" Then
                MsgAlarm = MsgAlarm & "¦C" & v & Chr(13)
                PtToLine = CInt(v)
            End If
        Next
        MsgBox MsgAlarm
        outSheet.Activate
        outSheet.Rows(PtToLine).Activate
    End If
   '==========¦¬§À==========
    DelSheet ("temp")
    Set inSheet = Nothing
    Set outSheet = Nothing
    Set tSheet = Nothing
    
End Sub

Private Sub Layer(Optional Mode As String = "DRCS")
' ****************************************************************************************************
' *
' * @Sub        Layer
' *
' * @Description
' *             To list layers used in boolean code and captured from worksheet "LayerTable"
' *
' ****************************************************************************************************

    Dim MAP As String
    Dim DATATYPE As String
    Dim iSheet As Worksheet
    Dim oSheet As Worksheet
    Dim tempStr As String
    Dim ErrorMsg As String
    Dim i As Long, j As Long, k As Long
    
    On Error GoTo Wrong
    
    Set iSheet = Worksheets("MES")
    Set oSheet = AddSheet("Layer Mapping")
        
    oSheet.Cells(1, 1).Value = Window
    flag = False
    ErrorMsg = "Cannot find following layers in LayerTable Sheet: " & Chr(13)
    
    For i = 1 To iSheet.UsedRange.Rows.Count
        If iSheet.Cells(i, 2).Value Like "*VARIABLE_START*" Then
            k = oSheet.UsedRange.Rows.Count
            For j = i + 1 To iSheet.UsedRange.Count
                If Worksheets("MES").Cells(j, 2).Value Like "*VARIABLE_END*" Then
                    Exit For
                Else
                    tempStr = iSheet.Cells(j, 2).Value
                    tempStr = Trim(tempStr)
                    tempStr = Replace(tempStr, " = a customized_data provided at tape out by customer", "")
                    tempStr = Replace(tempStr, "= a customized_data provided at tape out by customer", "")
                    tempStr = Replace(tempStr, " =a customized_data provided at tape out by customer", "")
                    tempStr = Replace(tempStr, "=a customized_data provided at tape out by customer", "")
                    tempStr = Replace(tempStr, " = a customized_data provided at tape out by customer ", "")
                    tempStr = Trim(tempStr)
                    oSheet.Cells(k + j - i, 1).Value = tempStr
                End If
            Next
        End If
    Next
    
        Columns(1).RemoveDuplicates Columns:=1, Header:=xlNo
    
    oSheet.Activate
    If Mode = "DRCS" Then
        For i = 1 To oSheet.UsedRange.Rows.Count
            MAP = Application.WorksheetFunction.VLookup(Cells(i, 1).Value, Worksheets("LayerTable").Range("A:C"), 2, 0)
            DATATYPE = Application.WorksheetFunction.VLookup(Cells(i, 1).Value, Worksheets("LayerTable").Range("A:C"), 3, 0)
              
            Cells(i, 1) = Replace(Cells(i, 1), "+", "PLUS")
            Cells(i, 1) = Cells(i, 1) & " = input(" & CStr(MAP) & ", " & CStr(DATATYPE) & ")"
            If IsNumeric(Left(Cells(i, 1), 1)) Then Cells(i, 1) = N2L(Left(Cells(i, 1), 1)) & Mid(Cells(i, 1), 2)
        Next i
    ElseIf Mode = "Calibre" Then
        For i = oSheet.UsedRange.Rows.Count To 2 Step -1
            oSheet.Rows(i).Insert
        Next i
        For i = 1 To oSheet.UsedRange.Rows.Count Step 2
            MAP = Application.WorksheetFunction.VLookup(Cells(i, 1).Value, Worksheets("LayerTable").Range("A:C"), 2, 0)
            DATATYPE = Application.WorksheetFunction.VLookup(Cells(i, 1).Value, Worksheets("LayerTable").Range("A:C"), 3, 0)
            Cells(i, 1) = "LAYER " & Replace(Cells(i, 1), "+", "PLUS")
            Cells(i + 1, 1) = "LAYER MAP " & MAP & "    DATATYPE " & DATATYPE
            Cells(i, 2) = 1000 + (i + 1) / 2
            Cells(i + 1, 2) = 1000 + (i + 1) / 2
        Next i
    End If
    
    Set iSheet = Nothing
    Set oSheet = Nothing

    If flag Then MsgBox ErrorMsg

    Exit Sub
Wrong:
    If Not InStr(ErrorMsg, Cells(i, 1).Value) > 0 Then ErrorMsg = ErrorMsg & Cells(i, 1).Value & Chr(13)
    flag = True
    MAP = ""
    DATATYPE = ""
    oSheet.Rows(i).Interior.ColorIndex = 3
    Resume Next

End Sub

Private Sub Result(Optional Mode As String = "DRCS")
' ****************************************************************************************************
' *
' * @Sub        Result
' *
' * @Description
' *             To print transform result on the final worksheet "Result".
' *
' ****************************************************************************************************
    
    If Mode = "Calibre" Then Call Result_Calibre: Exit Sub
    
    Dim iSheet As Worksheet
    Dim oSheet As Worksheet
    Dim dSheet As Worksheet
    Dim cSheet As Worksheet
    Dim i As Long
    
    Dim oRange As Range
    
    Dim DATATYPE As Integer
    Dim MAP As Integer
    Dim ErrorMsg As String
    
    Dim Layer() As String
    
    Set iSheet = Worksheets("Layer Mapping")
    Set oSheet = AddSheet("Result")
    Set dSheet = Worksheets(Mode)
    Set cSheet = Worksheets("CtrlTable")
    ErrorMsg = "Please define following output layers in LayerTable Sheet: " & Chr(13)
    
    iSheet.Activate
    iSheet.UsedRange.Copy
    oSheet.Activate
    oSheet.Cells(1, 1).Select
    oSheet.PasteSpecial
    dSheet.Activate
    dSheet.UsedRange.Copy
    oSheet.Activate
    oSheet.Cells(oSheet.UsedRange.Rows.Count + 2, 1).Select
    oSheet.PasteSpecial
    
    Set oRange = Range(Cells(oSheet.UsedRange.Rows.Count + 2, 1), Cells(oSheet.UsedRange.Rows.Count + cSheet.Cells(1, 1).CurrentRegion.Rows.Count + 1, 1))
    
    ReDim Layer(0) As String
    
    For i = 1 To oSheet.UsedRange.Rows.Count + 1
        If oSheet.Cells(i, 1).Value = "" Then
            Layer(UBound(Layer)) = getCOL(oSheet.Cells(i - 1, 1).Value, " = ", 1)
            ReDim Preserve Layer(UBound(Layer) + 1) As String
        End If
    Next i
    ReDim Preserve Layer(UBound(Layer) - 1) As String
    On Error GoTo Wrong
    For i = 1 To UBound(Layer)
        If cSheet.Cells(1, 1) = "" Or UCase(cSheet.Cells(1, 1)) = "END" Then Exit For
    
        MAP = Application.WorksheetFunction.VLookup(getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1), Worksheets("LayerTable").Range("A:C"), 2, 0)
        DATATYPE = Application.WorksheetFunction.VLookup(getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1), Worksheets("LayerTable").Range("A:C"), 3, 0)
    
        oRange.Cells(i, 1) = Layer(i) & ".output(" & CStr(MAP) & ", " & CStr(DATATYPE + subType) & ")"
    Next
     
    
    oSheet.Cells(1, 1).Select
    
    iSheet.Visible = xlSheetHidden
    dSheet.Visible = xlSheetHidden
    'Worksheets("MES").Visible = xlSheetHidden
    
    If flag Then MsgBox (Replace(Replace(ErrorMsg, "MINUS", "-"), "PLUS", "+")): oRange.Cells(1, 1).Select
    
    Set iSheet = Nothing
    Set oSheet = Nothing
    Set dSheet = Nothing

Exit Sub
Wrong:
    If Not InStr(ErrorMsg, getCOL(Layer(i), "_TEMP", 1)) > 0 Then ErrorMsg = ErrorMsg & getCOL(Layer(i), "_TEMP", 1) & Chr(13)
    flag = True
    MAP = -1
    oRange.Rows(i).Interior.ColorIndex = 3
    Resume Next


End Sub

Private Sub Result_Calibre()
' ****************************************************************************************************
' *
' * @Sub        Result
' *
' * @Description
' *             To print transform result on the final worksheet "Result".
' *
' ****************************************************************************************************
    
    Dim iSheet As Worksheet
    Dim oSheet As Worksheet
    Dim dSheet As Worksheet
    Dim cSheet As Worksheet
    Dim i As Long
    Dim k As Long
    
    Dim oRange As Range
    
    Dim DATATYPE As Integer
    Dim MAP As Integer
    Dim ErrorMsg As String
    Dim flagPt As Integer
    
    Dim Layer() As String
    
    Set iSheet = Worksheets("Layer Mapping")
    Set oSheet = AddSheet("Result")
    Set dSheet = Worksheets("Calibre")
    Set cSheet = Worksheets("CtrlTable")
    ErrorMsg = "Please define following output layers in LayerTable Sheet: " & Chr(13)
    
   '==========Title==========
            
    Cells(1, 1).Value = "//////////////////////////////////////////////////////   TITLE ////////////////////////////////////////////////////////////////"
    Cells(2, 1).Value = "PRECISION 1000"
    Cells(3, 1).Value = "LAYOUT MAGNIFY AUTO"
    Cells(4, 1).Value = "RESOLUTION 1"
    Cells(5, 1).Value = "LAYOUT INPUT EXCEPTION SEVERITY PRECISION_RULE_FILE 1"
    Cells(6, 1).Value = "LAYOUT ERROR ON INPUT YES"
    Cells(7, 1).Value = "DRC MAXIMUM RESULTS ALL"
    Cells(8, 1).Value = "DRC MAXIMUM VERTEX 199"
    Cells(9, 1).Value = "FLAG SKEW YES"
    Cells(10, 1).Value = "FLAG ACUTE YES"
    Cells(11, 1).Value = "FLAG OFFGRID YES"
    Cells(12, 1).Value = "LAYOUT PROCESS BOX RECORD YES"
    Cells(13, 1).Value = "DRC KEEP EMPTY NO"
    Cells(14, 1).Value = "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////"
    
    '==========Input Info==========
    
    Cells(18, 1).Value = "////////////////////////////////////////////////   INPUT INFO   ///////////////////////////////////////////////////////////"
    Cells(19, 1).Value = "LAYOUT SYSTEM GDSII"
    Cells(20, 1).Value = "LAYOUT PATH     ""[Layout Path]"""
    Cells(21, 1).Value = "LAYOUT PRIMARY  ""[Top Cell]"""
    Cells(22, 1).Value = ""
    Cells(23, 1).Value = "DRC RESULTS DATABASE    ""[Output db name]""  GDSII"
    Cells(24, 1).Value = "DRC SUMMARY REPORT      ""[Output summary report]"""
    Cells(25, 1).Value = "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////"

    '==========Layer Mapping==========
    
    Cells(29, 1).Value = "/////////////////////////////////////////////   LAYER MAPPING   //////////////////////////////////////////////////////"
    iSheet.Activate
    iSheet.UsedRange.Copy
    oSheet.Activate
    oSheet.Cells(30, 1).Select
    oSheet.PasteSpecial
    k = oSheet.UsedRange.Rows.Count
    Cells(k + 1, 1).Value = "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////"

    '==========Output Result==========

    Cells(k + 5, 1).Value = "/////////////////////////////////////////////   OUTPUT RESULT   //////////////////////////////////////////////////////"
    ReDim Layer(0) As String
    k = oSheet.UsedRange.Rows.Count + 1
    For i = 1 To dSheet.UsedRange.Rows.Count + 1
        If dSheet.Cells(i, 1).Value = "" Then
            Layer(UBound(Layer)) = getCOL(dSheet.Cells(i - 1, 1).Value, " = ", 1)
            ReDim Preserve Layer(UBound(Layer) + 1) As String
        End If
    Next i
    ReDim Preserve Layer(UBound(Layer) - 1) As String
    On Error GoTo Wrong
    For i = 1 To UBound(Layer) + 1
        If cSheet.Cells(1, 1) = "" Or UCase(cSheet.Cells(1, 1)) = "END" Then Exit For
    
        MAP = Application.WorksheetFunction.VLookup(getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1), Worksheets("LayerTable").Range("A:C"), 2, 0)
        DATATYPE = Application.WorksheetFunction.VLookup(getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1), Worksheets("LayerTable").Range("A:C"), 3, 0)

        oSheet.Cells(k + i, 1) = "DRC CHECK MAP  " & getCOL(Layer(i - 1), "_TEMP", 1) & "_NEW" & "         " & CStr(MAP) & " " & CStr(DATATYPE + subType) & "      " & getCOL(Layer(i - 1), "_TEMP", 1) & "_NEW" & "           {COPY " & Layer(i - 1) & "}"
    Next
    
    k = oSheet.UsedRange.Rows.Count
    oSheet.Cells(k + 2, 1) = "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////"
    oSheet.Cells(k + 6, 1) = "/////////////////////////////////////////////   BOOLEAN OPERATION   //////////////////////////////////////////////"
    
    
    '==========Boolean Operation==========
    
    dSheet.Activate
    dSheet.UsedRange.Copy
    oSheet.Activate
    oSheet.Cells(oSheet.UsedRange.Rows.Count + 2, 1).Select
    oSheet.PasteSpecial
    k = oSheet.UsedRange.Rows.Count
    
    Cells(1 + k, 1).Value = "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////"
        
    oSheet.Cells(1, 1).Select
    
    iSheet.Visible = xlSheetHidden
    dSheet.Visible = xlSheetHidden
    'Worksheets("MES").Visible = xlSheetHidden
    
    If flag Then MsgBox (ErrorMsg): oSheet.Cells(flagPt, 1).Select
    
    Set iSheet = Nothing
    Set oSheet = Nothing
    Set dSheet = Nothing

Exit Sub
Wrong:
    If Not InStr(ErrorMsg, getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1)) > 0 Then ErrorMsg = ErrorMsg & getCOL(cSheet.Cells(i, 1).Value, Chr(10), 1) & Chr(13)
    flag = True
    MAP = -1
    oSheet.Rows(k + i).Interior.ColorIndex = 3
    flagPt = k + i
    Resume Next


End Sub



Public Sub CombineRows()
' ****************************************************************************************************
' *
' * @Sub        CombineRows
' *
' * @Description
' *             To merge rows in worksheet "MES" into MT form format.
' *             It's a reverse step of "Initial"
' *
' ****************************************************************************************************
    
    Dim inSheet As Worksheet        'Sheet MES
    Dim outSheet As Worksheet       'Sheet NewTable
    Dim tempStr As String
    Dim tempHead As String
    Dim i As Long, j As Long        '°Ï°ìÅÜ¼Æ
    Dim LineCount As Integer
    
    '====================§PÂ_¦³¨S¦³MES====================
    If Not IsExistSheet("MES") Then MsgBox ("Sheet ""MES"" is not found."): Exit Sub
    
    AddSheet ("NewTable")
    
    Set inSheet = Worksheets("MES")
    Set outSheet = Worksheets("NewTable")
    
    LineCount = 0
    flag = False
    
    '====================±N³v¦æÂà´«¦¨MTFormªº®æ¦¡====================
    For i = 1 To inSheet.UsedRange.Rows.Count
        If Trim(inSheet.Cells(i, 1).Value) <> "" Then tempHead = inSheet.Cells(i, 1).Value
        If Trim(inSheet.Cells(i, 1).Value) = "" And Trim(inSheet.Cells(i, 2).Value) = "" Then
        Else
            If Trim(inSheet.Cells(i, 2).Value) <> "$VARIABLE_END" Then
                tempStr = tempStr & Chr(10) & TrimLn(inSheet.Cells(i, 2).Value)
            Else
                tempStr = tempStr & Chr(10) & "$VARIABLE_END"
                outSheet.Cells(1 + LineCount, 1).Value = tempHead
                outSheet.Cells(1 + LineCount, 2).Value = Mid(tempStr, Len(Chr(10)) + 1)
                LineCount = outSheet.UsedRange.Rows.Count
                tempStr = ""
            End If
        End If
    Next
    
    outSheet.Columns(1).AutoFit
    outSheet.Columns(2).ColumnWidth = 80
    
    Set inSheet = Nothing
    Set outSheet = Nothing

End Sub
Public Sub Scaling()
' ****************************************************************************************************
' *
' * @Sub        Scaling
' *
' * @Description
' *             To adjust scale of boolean operation due to scale ratio
' *             The operation is operated under worksheet "Result".
' *
' ****************************************************************************************************
    If Not IsExistSheet("Result") Then MsgBox "Sheet ""Result"" is not found.": Exit Sub
    
    Dim ratio
    Dim Pre
    Dim Post
    Dim digit As Integer
    Dim nowSheet As Worksheet
    Dim i As Long, j As Long
    Dim Val
    Dim tempAry
    Dim tempStr As String
    Dim nowRow As Integer
    Dim Mode As String

    nowRow = 1
    Mode = "DRCS"
    If Worksheets("Result").Cells(1, 1) = "//////////////////////////////////////////////////////   TITLE ////////////////////////////////////////////////////////////////" Then
        Mode = "Calibre"
        Do
            nowRow = nowRow + 1
        Loop Until Worksheets("Result").Cells(nowRow, 1) = "/////////////////////////////////////////////   BOOLEAN OPERATION   //////////////////////////////////////////////"
        nowRow = nowRow + 2
    End If
    ratio = Evaluate(InputBox("Please intput scale ratio:", "Scaling", "1/0.855"))
    Pre = InputBox("Please input boolean operation precision BEFORE scaling:", "Precision", 0.0002)
    Post = InputBox("Please input boolean operation precision AFTER scaling:", "Precision", 0.001)
    If Not IsNumeric(ratio) Then MsgBox "Please input effective number.": Exit Sub
    If Not IsNumeric(Pre) Then MsgBox "Please input effective number.": Exit Sub
    If Not IsNumeric(Post) Then MsgBox "Please input effective number.": Exit Sub
    digit = Abs(Log(Post) / Log(10))
    
    Set nowSheet = AddSheet("Result1")
    nowSheet.Columns(1).Value = Worksheets("Result").Columns(1).Value
    nowSheet.Columns(2).Value = Worksheets("Result").Columns(2).Value
    
    If Mode = "DRCS" Then
        For i = nowRow To nowSheet.UsedRange.Rows.Count
            If InStr(nowSheet.Cells(i, 1), ".sized") Then
                Val = getCOL(getCOL(nowSheet.Cells(i, 1), ".sized(", 2), ")", 1)
                If Not InStr(Val, ",") > 0 Then
                    If CDbl(Val) = Pre Or CDbl(Val) = -1 * Pre Then
                        tempStr = CStr(Post * Val / (Abs(Val)))
                    Else
                        tempStr = CStr(round(Val * ratio, digit))
                    End If
                Else
                    tempAry = Split(Val, ",")
                    For j = 0 To UBound(tempAry)
                        If CDbl(tempAry(j)) = Pre Or CDbl(tempAry(j)) = -1 * Pre Then
                            tempAry(j) = CStr(Post * Pre / (Abs(Pre)))
                        Else
                            tempAry(j) = CStr(round(tempAry(j) * ratio, digit))
                        End If
                    Next j
                    tempStr = Join(tempAry, ",")
                End If
    
                nowSheet.Cells(i, 1) = getCOL(nowSheet.Cells(i, 1), ".sized(", 1) & ".sized(" & tempStr & ")"
            ElseIf InStr(nowSheet.Cells(i, 1), "area") Then
                Val = getCOL(getCOL(nowSheet.Cells(i, 1), ".area", 2), "}", 1)
                Val = Replace(Val, ">", "")
                Val = Replace(Val, "<", "")
                Val = Replace(Val, "=", "")
                Val = CDbl(Trim(Val))
                tempStr = Replace(nowSheet.Cells(i, 1), Val, CStr(round(Val * ratio * ratio, digit)))
                nowSheet.Cells(i, 1) = tempStr
            ElseIf InStr(nowSheet.Cells(i, 1), "bbox") Then
                tempAry = Split(nowSheet.Cells(i, 1), " ")
                For j = 0 To UBound(tempAry)
                    If tempAry(j) = "polygon.bbox.width" Or tempAry(j) = "polygon.bbox.height" Then
                        If InStr(tempAry(j + 1), "=") Then
                            tempAry(j + 1) = Left(tempAry(j + 1), 2) & round(Mid(tempAry(j + 1), 3) * ratio, digit)
                        Else
                            tempAry(j + 1) = Left(tempAry(j + 1), 1) & round(Mid(tempAry(j + 1), 2) * ratio, digit)
                        End If
                    End If
                Next j
                nowSheet.Cells(i, 1) = Join(tempAry, " ")
            End If
        Next i
    ElseIf Mode = "Calibre" Then
        For i = nowRow To nowSheet.UsedRange.Rows.Count - 1
            tempAry = Split(nowSheet.Cells(i, 1), " ")
            For j = 0 To UBound(tempAry)
                If tempAry(j) = "area" Then flag = True
                If (IsNumeric(tempAry(j))) Then
                    If Abs(CDbl(tempAry(j))) = Pre Then
                        tempAry(j) = Post * (tempAry(j) / Abs(tempAry(j)))
                    Else
                        If flag Then tempAry(j) = tempAry(j) * ratio: flag = False
                        tempAry(j) = round(tempAry(j) * ratio, digit)
                    End If
                End If
            Next j
            tempStr = Join(tempAry, " ")
            flag = False
            
            nowSheet.Cells(i, 1) = tempStr
        Next i
    End If
    
    Set nowSheet = Nothing
    
End Sub

Private Sub Version()
    Ver.Show
End Sub

