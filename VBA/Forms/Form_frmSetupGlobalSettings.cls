VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetupGlobalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnBrowse_Click()
Dim fd As Office.FileDialog
Set fd = Access.FileDialog(msoFileDialogFolderPicker)
With fd
    .Filters.clear
    .title = "Select database ROOT location"
    If .show Then
        Dim s: For Each s In .SelectedItems
            dbRoot = s
        Next
    Else
        'Cancelled by user
    End If
End With
End Sub

Private Sub btnShiftDef_Click()
    DoCmd.OpenForm "frmShiftDef"
End Sub

Private Sub Form_AfterUpdate()
On Error Resume Next
    If Parent.Name = "frmBackOffice" Then
        Parent.Form!lblSave.Visible = Parent.Form.Dirty
    End If
End Sub

Private Sub Form_Dirty(cancel As Integer)
    If Parent.Name = "frmBackOffice" Then
        Parent.Form!lblSave.Visible = True
    End If
End Sub
