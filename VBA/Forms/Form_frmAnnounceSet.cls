VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAnnounceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub ctlEndEvent_BeforeUpdate(cancel As Integer)
cancel = False
    If Nz(ctlEndEvent) = "" Then Exit Sub
    If IsDate(ctlEndEvent) Then ctlEndEvent = Format(ctlEndEvent, "hhnn")
    If IsNumeric(ctlEndEvent) And Len(ctlEndEvent) = 3 Then ctlEndEvent = 0 & ctlEndEvent
    ctlEndEvent.BorderColor = RGB(166, 166, 166)
    Dim cTime As Variant: cTime = getTime4Char(ctlEndEvent)
    
    If Not IsDate(cTime) Then
        cancel = True
        ctlEndEvent = ""
        ctlEndEvent.BorderColor = RGB(255, 0, 0)
        Exit Sub
        'endEvent.SetFocus
    End If
    
    endEvent = Format(cTime, "hh:nn")
End Sub

Private Sub ctlStartEvent_BeforeUpdate(cancel As Integer)
cancel = False
    If Nz(ctlStartEvent) = "" Then Exit Sub
    If IsDate(ctlStartEvent) Then startEvent = Format(startEvent, "hhnn")
    If IsNumeric(ctlStartEvent) And Len(ctlStartEvent) = 3 Then ctlStartEvent = 0 & startEvent
    ctlStartEvent.BorderColor = RGB(166, 166, 166)
    Dim cTime As Variant: cTime = getTime4Char(ctlStartEvent)
    
    If Not IsDate(cTime) Then
        cancel = True
        ctlStartEvent = ""
        ctlStartEvent.BorderColor = RGB(255, 0, 0)
        Exit Sub
        'endEvent.SetFocus
    End If
    
    startEvent = Format(cTime, "hh:nn")
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub Form_Current()
    bClose = False
    ctlStartEvent = Format(startEvent, "hh:nn")
    ctlEndEvent = Format(endEvent, "hh:nn")
End Sub

Private Sub dom_Click()
    rrefresh
End Sub

Private Sub done_Click()
bClose = True
    If IsNull(startDate) Then
        MsgBox "Please enter a Start Date.", vbInformation, "AeroStat"
        Exit Sub
    End If
    If IsNull(msg) Then
        MsgBox "Please enter a Message.", vbInformation, "AeroStat"
        Exit Sub
    End If
    
    sscope = GetNum(sscope)
    
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close acForm, "frmAnnounceSet"
    
End Sub

Private Sub endDate_LostFocus()
rrefresh
End Sub

Private Sub Form_Load()
    rrefresh
End Sub

Private Sub recur_Click()
    rrefresh
End Sub

Private Sub scope_LostFocus()
    rrefresh
End Sub

Private Function rrefresh()
    Select Case recur
        Case 1 'Never
            For Each ctl In Controls
                If Not ctl.Tag = "" Then ctl.Visible = False
            Next
        
        Case 2 'Daily
            lblScope.Caption = "Day" & IIf(scope > 1, "s", "")
        
            For Each ctl In Controls
                If ctl.Tag Like "w" Then
                    ctl.Visible = True
                ElseIf ctl.Tag Like "m" Then
                    ctl.Visible = False
                End If
            Next
        
        Case 3 'Weekly
            lblScope.Caption = "Week" & IIf(scope > 1, "s", "")
        
            For Each ctl In Controls
                If ctl.Tag Like "w" Then
                    ctl.Visible = True
                ElseIf ctl.Tag Like "m" Then
                    ctl.Visible = False
                End If
            Next
        
        Case 4 'Monthly
            lblScope.Caption = "Month" & IIf(scope > 1, "s", "")
            
            If Not dom Then
                lblSScope.Caption = Format(Weekday(startDate), "dddd")
            Else
                lblSScope.Caption = "day"
            End If
            
            For Each ctl In Controls
                If ctl.Tag Like "w" Or ctl.Tag Like "m" Then ctl.Visible = True
            Next
        
        Case 5 'Yearly
        'Something unique
        
    End Select
    
    sscope = numTense(Nz(sscope))
    
    If recur = 1 Then
        scopeText = "Only once"
    Else
        scopeText = "Repeat every " & scope & " " & lblScope.Caption & IIf(recur = 4, " on the " & sscope & " " & lblSScope.Caption, "")
    End If

End Function

Private Sub sscope_LostFocus()
    rrefresh
End Sub

Private Sub startDate_AfterUpdate()
endDate = Nz(startDate, endDate)
End Sub

Private Sub startDate_LostFocus()
    rrefresh
End Sub

