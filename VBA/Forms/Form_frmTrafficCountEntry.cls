VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficCountEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub acount_Click()
    autocount
End Sub

Private Sub done_Click()
On Error Resume Next
    If Me.NewRecord And Not CurrentDb.OpenRecordset("SELECT * FROM tblAnalytics WHERE Format(tblAnalytics.countDay,'dd-mmm-yy') = '" & Format(Me.countDay, "dd-mmm-yy") & "'").RecordCount = 0 Then
        MsgBox "This day should already exist. If this is an error, please report it using the feedback button.", vbInformation, "AeroStat"
        Exit Sub
    End If
    
    bClose = True

    For Each ctl In Controls
        If TypeOf ctl Is TextBox And Not ctl.Name = "countDay" Then
            If ctl = "" Then ctl = 0
        End If
    Next
    
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close acForm, "frmTrafficCountEntry", acSaveYes
    If Forms!frmMain!view.SourceObject = "frmTrafficCount" Then
        Forms!frmMain!view!shiftCount.Requery
        Forms!frmMain!view!shiftCountTotal.Requery
    End If
    
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
    bClose = False
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub Form_Load()
    If Me.NewRecord Then
        countDay.Enabled = True
        countDay = Date
        autocount True
        MsgBox "Some totals have been automatically generated." & vbCrLf & "Please ensure their accuracy.", vbInformation, "AeroStat"
    End If
End Sub

Private Function autocount(Optional ByVal newrec As Boolean)
    If newrec Then
        If hazCargo = 0 Then hazCargo = Abs(DSum("hazCargo", "tblTraffic", "Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format(tblAnalytics.countDay,'mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')"))
        If DVs = 0 Then DVs = DCount("dvCode", "tblTraffic", "Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format('" & countDay & "','mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')")
        If transientArr = 0 Then transientArr = Abs(DSum("transient", "tblTraffic", "(Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format('" & countDay & "','mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')) And Status = 'Closed'"))
    Else
        If MsgBox("Haz Cargo, DVs, and Transient counts will be reset." & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "AeroStat") = vbYes Then
            hazCargo = Abs(DSum("hazCargo", "tblTraffic", "Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format('" & countDay & "','mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')"))
            DVs = DCount("dvCode", "tblTraffic", "Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format('" & countDay & "','mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')")
            transientArr = Abs(DSum("transient", "tblTraffic", "(Format('" & countDay & "','mm/dd/yyyy') = Format(arrDate,'mm/dd/yyyy') Or Format('" & countDay & "','mm/dd/yyyy') = Format(depDate,'mm/dd/yyyy')) And Status = 'Closed'"))
        End If
    End If
End Function
