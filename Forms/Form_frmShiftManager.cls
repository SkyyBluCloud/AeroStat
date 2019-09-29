VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public dashShiftID As Integer
'SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart, tblShiftManager.shiftEnd, tblShiftManager.shiftTitle, IIf([closed],"Closed","Open") AS Status, tblShiftManager.superLead, tblShiftManager.amoc1, tblShiftManager.amoc2, tblShiftManager.amoc3, tblShiftManager.amoc4
'FROM tblShiftManager
'ORDER BY IIf([closed],"Closed","Open") DESC , tblShiftManager.shiftStart DESC , tblShiftManager.shiftEnd DESC;

Private Sub back_Click()
changeView "frmShiftList"
End Sub

Private Sub Command11_Click()
changeView "frmAFMDash"
End Sub

Private Sub Form_Load()
shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
dashShiftID = IIf(isClosed, 0, shiftID)
If dashShiftID <> 0 Then
    changeView "frmShiftDetails"
    view.Form!sf3616.SourceObject = "frm3616"
End If
'Me.view.sfshiftbrief.Form.filter = "shiftID = " & dashShiftID
'Me.view.sfshiftbrief.Form.FilterOn = True
End Sub

Private Sub newShift_Click()
On Error Resume Next
shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)

    If Not isClosed Then
        MsgBox "To create a new shift, the last one must be closed first.", vbInformation, "Shift Manager"
'        changeView "frmShiftDetails"
'        With view.Form!sfshiftbrief.Form
'            .filter = "shiftID = " & shiftID
'            .FilterOn = True
'        End With
'        Exit Sub
    End If

changeView "frmShiftDetails"
End Sub

Private Sub return_Click()
DoCmd.Close acForm, "frmShiftManager", acSaveYes
End Sub

Private Sub shiftView_DblClick(cancel As Integer)
    DoCmd.OpenForm "frmShiftNew", , , "shiftID = " & shiftView.Column(0), acFormEdit
End Sub

Public Function changeView(ByVal v As String, Optional ByVal Shift As Integer)
    view.SetFocus
    back.Visible = (v <> "frmShiftList")
    If Shift <> 0 Then dashShiftID = Shift
    view.SourceObject = v
    
'    With view.Form!sfshiftbrief.Form
'        .filter = "shiftID = " & shiftID
'        .FilterOn = True
'    End With
End Function

Public Function getShiftID() As Integer

    If dashShiftID <> 0 Then
        getShiftID = dashShiftID
        Exit Function
    End If
    
    shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
    isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
    getShiftID = IIf(isClosed, 0, shiftID)
    'frm!filter = "shiftID = " & dashShiftID
    'frm!FilterOn = True
End Function

