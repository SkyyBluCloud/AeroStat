VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CONSOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function command(ByVal cmd As String) As Boolean
command = True
Dim args() As String
args = Split(cmd, " ")

Select Case args(0)
    Case "/help" 'Table of Contents
        log "/table - opens table", "CONSOLE"
        log "/eval - Run Eval(exp)", "CONSOLE"
        
    Case "/table"
        DoCmd.OpenTable args(1)
        
    Case "/eval"
        log CStr(Eval(args(1))), "CONSOLE"
        
    Case Else
        log "Invalid command " & args(0), "CONSOLE", "ERR"
        command = False
        
End Select
        
End Function

Private Sub ctlInput_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
    Select Case KeyCode
        Case vbKeyReturn
            If Left(ctlInput.Text, 1) = "/" Then
                command ctlInput.Text
            Else
                log Eval("""" & ctlInput.Text & """"), "CONSOLE"
            End If
        Case Else
            Exit Sub
    End Select
    
sexit:
    ctlInput = ""
    update
    
    Exit Sub
errtrap:
    ErrHandler err, Error$, "CONSOLE"
    GoTo sexit
End Sub

Public Function update()
    lstFeed.Requery
    lstFeed.Selected(lstFeed.ListCount - 1) = True
    lstFeed.Selected(lstFeed.ListCount - 1) = False
    ctlInput.SetFocus
End Function

Private Sub Form_Load()
update
End Sub
