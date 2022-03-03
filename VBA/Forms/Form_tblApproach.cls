VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_tblApproach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim bClose As Boolean

Private Sub btnSave_Click()
If Not isValid Then Exit Sub
    
    bClose = True
    If Dirty Then Dirty = False
    DoCmd.Close acForm, Me.Name
    
End Sub

Private Sub ctlETA_BeforeUpdate(cancel As Integer)
If Nz(ctlETA) = "" Then Exit Sub
    If IsDate(ctlETA) Then ctlETA = Format(ctlETA, "hhnn")
    If IsNumeric(ctlETA) And Len(ctlETA) = 3 Then ctlETA = 0 & ctlETA
    ctlETA.BorderColor = RGB(166, 166, 166)
    ETA = getTime4Char(ctlETA)
    ctlETA = Left(Format(ctlETA, "0000"), 2) & ":" & Right(Format(ctlETA, "0000"), 2)
    
    If Not IsDate(ctlETA) Then
        'cancel = True
        ctlETA = ""
        ctlETA.BorderColor = RGB(255, 0, 0)
        ctlETA.SetFocus
    End If
End Sub

Private Function isValid() As Boolean
isValid = True

    Dim ctl: For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ListBox Then
        
            If IsNull(ctl.Value) Then
                ctl.BorderColor = vbRed
                isValid = False
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next
End Function

Private Sub Form_BeforeUpdate(cancel As Integer)
If Not bClose Then
    Undo
    Exit Sub
End If
End Sub

Private Sub Form_Current()
bClose = False
End Sub
