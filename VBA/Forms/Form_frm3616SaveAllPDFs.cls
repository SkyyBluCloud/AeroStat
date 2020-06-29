VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm3616SaveAllPDFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub save_Click()
    
    If Nz(endDate) = "" Then
        endDate = Date
    End If
    
    If Nz(startDate) = "" Then
        startDate.BorderColor = RGB(255, 0, 0)
        Exit Sub
    Else
        startDate.BorderColor = RGB(166, 166, 166)
    End If
    If MsgBox("This process might take a while. Continue?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub

    Status.Caption = "Saving. Please wait...."
    Status.Visible = True
    DoEvents
    UtilForm3616.saveAllPDFs startDate, endDate, Me.optOverwrite
    Status.Caption = "Success!"
    DoEvents
    MsgBox "Save successful!", vbInformation, "AF3616 to PDF"
    DoCmd.Close acForm, Me.Name
End Sub
