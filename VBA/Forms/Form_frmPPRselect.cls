VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPPRselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnAddToBaseAcft_Click()
    If MsgBox("Add aircraft to inventory?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
        DoCmd.OpenForm "frmBaseAcftAdd", , , , acFormAdd, acDialog, Forms!quick_input!Type & ";" & Forms!quick_input!Tail
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub custom_Click()
    With Forms!quick_input
        !PPR = InputBox("Please enter a PPR number:", "PPR", "-")
        If IsNull(DLookup("PPR", "tblppr", "ppr = '" & !PPR & "'")) Then
            MsgBox "This PPR doesn't exist. Please create a new one.", vbInformation, "Flight Plan"
            !PPR = ""
            Exit Sub
        End If
    End With
    
    DoCmd.Close acForm, "frmPPRselect"
End Sub

Private Sub newPPR_Click()
    DoCmd.OpenForm "new_ppr", , , , acFormAdd, acDialog, Me.Name
    pprList.Requery
End Sub

Private Sub save_Click()
If IsNull(pprList.Column(2)) Then Exit Sub

    With Forms!quick_input
        !PPR = pprList.Column(2)
    End With
    
    DoCmd.Close acForm, "frmPPRselect"
End Sub
