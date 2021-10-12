VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStereoFlightPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bClose As Boolean

Private Sub btnDelete_Click()
    If MsgBox("This flight plan will be deleted.", vbInformation + vbOKCancel, "Stereo Manager") = vbOK Then
        log "Stereo plan " & Tag & " deleted.", Me.Name & "btnDelete_click"
    End If
End Sub

Private Sub btnSave_Click()

    If Dirty Then
        bClose = True
        Dirty = False
        log "Stereo plan " & Tag & " updated.", Me.Name & "btnSave_click"
    End If

    lstMain.Requery
    tabCtl = 0
End Sub

Private Sub ctlETE_Exit(cancel As Integer)
    If Nz(ctlETE) = "" Then Exit Sub
    If IsDate(ctlETE) Then ctlETE = Format(ctlETE, "hhnn")
    If IsNumeric(ctlETE) And Len(ctlETE) = 3 Then ctlETE = 0 & ctlETE
    ctlETE.BorderColor = RGB(166, 166, 166)
    ETE = getTime4Char(ctlETE)
    ctlETE = Left(Format(ctlETE, "0000"), 2) & ":" & Right(Format(ctlETE, "0000"), 2)
    
    If Not IsDate(ctlETE) Then
        'cancel = True
        ctlETE = ""
        ctlETE.BorderColor = RGB(255, 0, 0)
        ctlETE.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If Not CurrentProject.Name Like "*DEV*" Then tabCtl.Style = 2 'TabStyle "None"
End Sub

Private Sub btnBack_Click()
    lstMain.SetFocus
End Sub

Private Sub Form_Current()
    bClose = False
    If Not IsNull(ETE) Then ctlETE = Hour(ETE) & ":" & Format(Minute(ETE), "00")
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        'cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub btnNew_Click()
    Me.DataEntry = True
    stereoTag.SetFocus
End Sub

Private Sub lstMain_DblClick(cancel As Integer)
    If Not IsNull(lstMain) Then
        With Me
            .DataEntry = False
            .Filter = "ID = " & lstMain
            .FilterOn = True
            .tabCtl = 1
        End With
    End If
End Sub

Private Sub TabCtl_Change()
    btnBack.Visible = (tabCtl <> 0)
    If tabCtl = 0 Then Me.FilterOn = False
End Sub
