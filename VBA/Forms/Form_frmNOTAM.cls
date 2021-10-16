VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNOTAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub afi11208_Click()
DoCmd.OpenForm "regViewer", , , , , , "https://static.e-publishing.af.mil/production/1/af_a3/publication/afi11-208/afi11-208.pdf"
'https://static.e-publishing.af.mil/production/1/af_a3/publication/afi11-208/afi11-208.pdf
End Sub

Private Sub btnImport_Click()
On Error GoTo errtrap
Dim fd As Office.FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
Dim f As String

    If MsgBox("The NOTAM report should begin at the earliest possible time, and end on the current date. " & _
                "Doing this will replace existing data in the table. Use this only if you know what you are doing.", _
                vbCritical + vbOKCancel, "NOTAM Report") = vbCancel Then Exit Sub
    
    With fd
        .title = "Select NOTAM report"
        .Filters.clear
        .Filters.add "Comma Separated Values", "*.csv"
        
        If .show Then
            Dim i: For Each i In .SelectedItems
                f = i.Value
            Next
            
            Util.trunc "tblNOTAM"
            DoCmd.TransferText , , "tblNOTAM", f, True
            log "NOTAM Import Successful", "frmNOTAM.btnImport_Click"
            view.SourceObject = "frmNOTAMview"
        Else
            GoTo errtrap
        End If
    End With

    
sexit:
    
    Exit Sub
errtrap:
    Select Case err
    
    End Select
End Sub
