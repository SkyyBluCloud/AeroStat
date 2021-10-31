VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRestrictions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'OpenArgs = The date to be checked
Private Sub cancel_Click()
    Forms!new_ppr.frmRestrictionsResponse = True
    DoCmd.Close acForm, "frmrestrictions", acSaveNo
End Sub

Private Sub continue_Click()
DoCmd.Close acForm, "frmrestrictions", acSaveNo
End Sub

Private Sub Form_Load()
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs("qAnnounce")
    qdf.Parameters("cd1") = Split(OpenArgs, ";")(0)
    qdf.Parameters("cd2") = Split(OpenArgs, ";")(1)
    Set restrictions.Recordset = qdf.OpenRecordset
End Sub
