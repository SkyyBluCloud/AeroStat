VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub add_Click()
    DoCmd.OpenForm "frmAnnounceSet", , , , acFormAdd
End Sub

Private Sub announcements_DblClick(cancel As Integer)
    DoCmd.OpenForm "frmAnnounceSet", , , "ID=" & announcements.Column(0), acFormEdit
End Sub

Private Sub announcements_GotFocus()
    announcements.Requery
End Sub

Private Sub del_Click()
If IsNull(Me.announcements) Then Exit Sub
    If MsgBox("Delete notice?", vbQuestion + vbYesNo, "Notice") = vbYes Then
        CurrentDb.Execute "DELETE FROM tblAnnounce WHERE ID = " & Me.announcements
        Me.announcements.Requery
    End If
End Sub

'SELECT tblAnnounce.ID, IIf(startDate=endDate,startDate,startDate & ' - ' & endDate) AS [Date Range], IIf((startEvent & ' - ' & endEvent)=' - ','All Day',startEvent & ' - ' & endEvent) AS [Time Range], tblAnnounce.msg
'From tblAnnounce
'WHERE (((tblAnnounce.startDate)=[checkDate]) AND ((serialdate([checkDate])) Between serialDate([startDate]) And serialDate([endDate]))) OR (((serialdate([checkDate])) Between serialDate([startDate]) And serialDate([endDate])) AND ((tblAnnounce.recur)=2)) OR (((serialdate([checkDate])) Between serialDate([startDate]) And serialDate([endDate])) AND ((tblAnnounce.recur)=3) AND ((DateDiff('d',[startDate],[checkDate]) Mod ([scope]*7))=0)) OR (((serialdate([checkDate])) Between serialDate([startDate]) And serialDate([endDate])) AND ((tblAnnounce.recur)=4) AND ((tblAnnounce.dom)=True) AND ((Day([subscope]))=Day([checkDate])) AND ((DateDiff('m',[startDate],[checkDate]) Mod [scope])=0)) OR (((serialdate([checkDate])) Between serialDate([startDate]) And serialDate([endDate])) AND ((tblAnnounce.recur)=4) AND ((tblAnnounce.dom)=False) AND ((DateDiff('m',[startDate],[checkDate]) Mod [scope])=0) AND ((DateDiff('d',DateAdd('d',IIf(Weekday([startDate])-Weekday(DateSerial(Year([checkDate]),Month([checkDate]),1))<0,7-Abs(W
'eekday([startDate])-Weekday(DateSerial(Year([checkDate]),Month([checkDate]),1))),Weekday([startDate])-Weekday(DateSerial(Year([checkDate]),Month([checkDate]),1))),DateSerial(Year([checkDate]),Month([checkDate]),1)),[checkDate]))=7*([subScope]-1)));
