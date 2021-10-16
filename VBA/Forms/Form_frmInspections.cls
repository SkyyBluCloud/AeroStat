VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInspections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private fail As Boolean
Private bClose As Boolean
Public varDate As Date

Private Sub btn483_Click()
    DoCmd.OpenForm "frm483compcheck", , , , acFormAdd, acDialog, Parent!shiftID
End Sub

Private Sub btnBack_Click()
tabCtl = btnBack.Tag
End Sub

Private Sub btnNewInsp_Click()
Me.DataEntry = True
tabCtl.Pages(1).SetFocus
btnBack.Tag = 0
End Sub

Private Sub btnSave_Click()
On Error GoTo errtrap
    
    fail = False
    Dim c: For Each c In Me.Controls
        If c.Tag = "required" Then
            If TypeOf c Is TextBox Or TypeOf c Is ComboBox Then
                If Nz(c.Value, 0) = 0 Then
                    c.BorderColor = RGB(255, 0, 0)
                    fail = True
                Else
                    c.BorderColor = RGB(166, 166, 166)
                End If
            End If
        End If
    Next c
    If fail Then Exit Sub
    
    bClose = True
    If Me.Dirty Then
        Me.Dirty = False
        With lstInsp
            .Requery
            .SetFocus
        End With
    End If
    
    Call lstInspUpdate
    
    

sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".btnSave_Click"
End Sub

Private Sub lstInspUpdate()
    Dim newSQL As String: newSQL = Replace(CurrentDb.QueryDefs("qInspectionsShiftView").sql, "[varDate]", "#" & DateValue(Nz(Parent!shiftStart, Date)) & "#")
    Me.lstInsp.RowSource = newSQL
End Sub

Private Sub ctlInTime_AfterUpdate(): With ctlInTime
    
    Dim ct As Variant: ct = getTime4Char(ctlInTime)
    If Not IsDate(ct) Then
        .BorderColor = vbRed
    Else
        .BorderColor = RGB(166, 166, 166)
        .Value = ct
    End If
End With
End Sub

Private Sub ctlOutTime_AfterUpdate(): With ctlOutTime

    Dim ct As Variant: ct = getTime4Char(ctlOutTime)
    If Not IsDate(ct) Then
        .BorderColor = vbRed
    Else
        .BorderColor = RGB(166, 166, 166)
        .Value = ct
    End If
End With
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then Undo
End Sub

Private Sub Form_Current()
    bClose = False
    
    Dim c: For Each c In Me.Controls
        If c.Tag Like "*!newrecord*" Then
            c.Visible = Not NewRecord
        End If
    Next c
    
End Sub

Private Sub Form_Load()
    If Not CurrentProject.Name Like "*DEV*" Then tabCtl.Style = 2  'No tabs
    Call lstInspUpdate
End Sub

Private Sub lstInsp_DblClick(cancel As Integer)
If IsNull(lstInsp) Then Exit Sub
On Error GoTo errtrap
    
    Me.filter = "tblInspections.ID = " & lstInsp
    Me.FilterOn = True
    tabCtl.Pages(1).SetFocus
    btnBack.Tag = 0
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".lstInsp_DblClick"
End Sub

Private Sub TabCtl_Change()
tabCtl.Pages(tabCtl).SetFocus
btnBack.Visible = tabCtl <> 0
End Sub
