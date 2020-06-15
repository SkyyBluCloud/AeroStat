VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInspections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public shiftID As Integer

Public Function changeView(ByVal v As String, Optional ByVal sid As Integer, Optional ByVal filter As Variant = 0)
    shiftID = sid
    view.SourceObject = v
    If filter <> 0 Then
        view.Form.Undo
        view.Form.filter = "ID = " & filter
        view.Form.FilterOn = True
    End If
'    If v = "frmInspectionNew" Then
'    With view.Form
'        If IsNull(!outTime) Then
'            !outTime = LToZ(Format(Now, "hh:nn"))
'        ElseIf IsNull(!inTime) Then
'            !inTime = LToZ(Format(Now, "hh:nn"))
'        End If
'    End With
'    End If
End Function

Private Sub back_Click()
    changeView back.Tag
    view.SetFocus
    back.Visible = False
End Sub
