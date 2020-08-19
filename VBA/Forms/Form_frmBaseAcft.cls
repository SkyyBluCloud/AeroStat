VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBaseAcft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub acType_Change()
acType.ForeColor = RGB(0, 0, 0)
acType.FontItalic = False
If Nz(acType.Value) = Replace(acType.DefaultValue, """", "") Then
    acType = acType.Text
    acType.SelStart = Len(acType)
End If
End Sub

Private Sub acType_Enter()
If Nz(acType.Value) = Replace(acType.DefaultValue, """", "") Then acType = ""

End Sub

Private Sub acType_Exit(cancel As Integer)
doCAPS Me
replaceGhost
End Sub

Private Sub btnBaseAcftAdd_Click()
Me.DataEntry = True
Me.Tail.SetFocus
End Sub

Private Sub btnSave_Click()
On Error GoTo errtrap
bClose = True
Dim ctl: For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
        If ctl.Tag = "ghost" Then
            If ctl.Value = Replace(ctl.DefaultValue, """", "") Then ctl.Value = ""
        End If
    End If
Next
sexit:
    If Dirty Then Dirty = False
    lstBaseAcft.Requery
    DataEntry = True
    Tail.SetFocus
    acType.BorderColor = RGB(166, 166, 166)
errtrap:
    Select Case err
    Case 3201
        acType.BorderColor = RGB(255, 0, 0)
    End Select
End Sub

Private Sub designator_Change()
designator.ForeColor = RGB(0, 0, 0)
designator.FontItalic = False
If Nz(designator.Value) = Replace(designator.DefaultValue, """", "") Then
    designator = designator.Text
    designator.SelStart = Len(designator)
End If
End Sub

Private Sub designator_Enter()
If Nz(designator.Value) = Replace(designator.DefaultValue, """", "") Then designator = ""

End Sub

Private Sub designator_Exit(cancel As Integer)
replaceGhost
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
If Not bClose Then
    'Cancel = True
    Undo
    Exit Sub
End If
End Sub

Private Sub Form_Current()
bClose = False
lblSave.Visible = Dirty
replaceGhost
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
keySave KeyCode
End Sub

Public Sub keySave(ByVal KeyCode As Integer)
Select Case KeyCode
    Case VBA.vbKeyReturn
        btnSave.SetFocus
        btnSave_Click
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyCode
    Case VBA.vbKeyReturn
        btnSave.SetFocus
        btnSave_Click
End Select
End Sub

Private Sub lstBaseAcft_AfterUpdate()
    Me.filter = "ID = " & lstBaseAcft
    Me.FilterOn = True
    'replaceGhost
End Sub

Public Function replaceGhost(Optional ByRef ctlTxt As TextBox, Optional ByVal enter As Boolean)
'    If enter Then
'        txt.ForeColor = RGB(0, 0, 0)
'        txt = ""
'        txt.FontItalic = False
'    ElseIf Nz(txt) = "" Then
'        txt.ForeColor = RGB(166, 166, 166)
'        txt = txt.OldValue
'        txt.FontItalic = True
'    End If
DoEvents

    Dim ctl: For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Tag = "ghost" Then
                If Nz(ctl.Value) <> "" And Nz(ctl.Value) <> Replace(ctl.DefaultValue, """", "") Then
                    ctl.ForeColor = RGB(0, 0, 0)
                    ctl.FontItalic = False
                Else
                    ctl.ForeColor = RGB(166, 166, 166)
                    ctl.Value = Replace(ctl.DefaultValue, """", "")
                    ctl.FontItalic = True
                End If
            End If
        End If
    Next
        
    
End Function

Private Sub lstBaseAcft_Click()
On Error Resume Next
    Me.filter = "ID = " & lstBaseAcft
    Me.FilterOn = True
End Sub

Private Sub Tail_Change()
Tail.ForeColor = RGB(0, 0, 0)
Tail.FontItalic = False
If Nz(Tail.Value) = Replace(Tail.DefaultValue, """", "") Then
    Tail = Tail.Text
    Tail.SelStart = Len(Tail)
End If
End Sub

Private Sub Tail_Enter()
If Nz(Tail.Value) = Replace(Tail.DefaultValue, """", "") Then Tail = ""
End Sub

Private Sub Tail_Exit(cancel As Integer)
doCAPS Me
replaceGhost
End Sub

Private Sub Unit_Enter()
If Nz(Unit.Value) = Replace(Unit.DefaultValue, """", "") Then Unit = ""
End Sub
