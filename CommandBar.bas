Option Explicit

Private Sub Workbook_Open()

Dim CB As CommandBar
Dim CBCtrl As CommandBarControl
Dim CBB As CommandBarButton

    Application.ScreenUpdating = False

Set CB = Nothing

On Error Resume Next
    
    Application.CommandBars("U2C").Delete

On Error GoTo 0

Set CB = Application.CommandBars.Add(Name:="U2C", Temporary:=False)

'Button


Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Initial"
    .FaceId = 601
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Initial"
    .Enabled = True
End With

Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Execute(KLayout)"
    .FaceId = 136
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!AutoRun"
    .Enabled = True
End With

Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Execute(Calibre)"
    .FaceId = 136
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!AutoRun_Calibre"
    .Enabled = True
End With

Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "MergeRows"
    .FaceId = 37
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!CombineRows"
    .Enabled = True
End With

Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Scaling"
    .FaceId = 966
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Scaling"
    .Enabled = True
End With

Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Ver. 1.13"
    .FaceId = 487
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Version"
    .Enabled = True
End With

'Menu
With Application.CommandBars("DRCS")
    .Visible = True
    .Position = msoBarTop
End With

    Application.ScreenUpdating = True

End Sub

Private Sub Workbook_Deactivate()

On Error Resume Next
    
    Application.CommandBars("U2C").Visible = False

On Error GoTo 0

End Sub

Private Sub Workbook_activate()

On Error Resume Next
    
    Application.CommandBars("U2C").Visible = True

On Error GoTo 0

End Sub




