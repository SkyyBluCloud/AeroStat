VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function afldStatusUpdate(ByRef fld As Control)
On Error GoTo errtrap
If DLookup("closed", "tblshiftmanager", "shiftid = " & DLookup("lastshift", "tbluserauth", "username = '" & getUser & "'")) Then GoTo fexit

    Select Case fld.Name
        Case "rwy"
            If fld.OldValue <> fld Then barrier = "DOWN"
            
        Case "rsc"
            Select Case fld.Column(1)
            Case True
                Do
                    Me.rcr = InputBox("What is the RCR?", "RCR", "Ex: 12")
                    
                Loop While Not IsNull(rcr) And Not IsNumeric(Me.rcr)
            
            Case Else
                Me.rcr = Null
            End Select
            fld = Me.rsc & Me.rcr
            
        Case "bwc"
        
        Case "barrier"
            
        
        Case Else
            Exit Function
            
    End Select
    
    
fexit:
    If fld.OldValue <> fld Then
        If Dirty Then Dirty = False
        FormSyncUtil.syncForm "frmTrafficLog"
        FormSyncUtil.syncForm "frmFlightMonitor"
    End If
    
    focus.SetFocus
    Exit Function
errtrap:
    ErrHandler err, Error$, Me.Name & ".afldStatusUpdate(" & fld & ")"
End Function


Private Sub Form_Load()
    If DLookup("authlevel", "tblUserAuth", "username = '" & getUser & "'") > 7 Then
        Dim ctl: For Each ctl In Controls
            If TypeOf ctl Is ComboBox Then
                ctl.Locked = True
            End If
        Next ctl
    End If
End Sub

Private Sub Form_Timer()
    Dim ctl: For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            With ctl
                '.enabled = True
                .Requery
                FilterOn = True
                DoEvents
                '.enabled = False
            End With
        End If
    Next ctl
    focus.SetFocus
    DoEvents
End Sub
