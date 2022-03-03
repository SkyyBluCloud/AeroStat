VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBaseAcftAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub acType_Enter()
replaceGhost Me.acType, True
End Sub

Private Sub acType_Exit(cancel As Integer)
replaceGhost Me.acType
End Sub

Private Sub btnUpdate_Click()
    DoCmd.Close acForm, "frmBaseAcftAdd"
End Sub

Private Sub designator_Enter()
replaceGhost Me.designator, True
End Sub

Private Sub designator_Exit(cancel As Integer)
replaceGhost Me.designator
End Sub

Private Sub Form_Load()

'        With Parent.Form!lstBaseAcft
'            .Value = DLookup("ID", "tblbaseacft")
'            .MoveFirst
'            Dim a: a = .Column(2) & ";" & .Column(1)
'        End With
'    End If
    If Not IsNull(openargs) Then
        Dim s() As String
        s = Split(openargs, ";")
        acType = s(0)
        Tail = s(1)
    End If
End Sub

Private Sub Tail_Enter()
    replaceGhost Me.Tail, True
End Sub

Public Function replaceGhost(ByRef txt As TextBox, Optional ByVal enter As Boolean)
    If enter Then
        txt.ForeColor = RGB(0, 0, 0)
        txt = ""
        txt.FontItalic = False
    ElseIf Nz(txt) = "" Then
        txt.ForeColor = RGB(166, 166, 166)
        txt = txt.OldValue
        txt.FontItalic = True
    End If
End Function

Private Sub Tail_Exit(cancel As Integer)
    replaceGhost Me.Tail
End Sub
