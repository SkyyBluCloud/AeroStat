VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmParkingMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
On Error GoTo errtrap
Dim RS As DAO.Recordset
Set qdf = CurrentDb.QueryDefs("qOnStation")
qdf.Parameters("bos") = True
Set RS = qdf.OpenRecordset
With RS
Do While Not .EOF
    If Not Nz(!Spot) = "" Then
        
        If InStr(1, !Spot, ",") > 0 Then
            Dim r As String 'ramp
            Dim N As New Collection 'number
            Dim sspot As String
            sspot = Replace(!Spot, ",", "")
            
            For i = 1 To Len(sspot)
                char = Mid(sspot, i, 1)
                If IsNumeric(char) Then
                    N.add i
                Else
                    r = r & char
                End If
            Next
            
            For Each i In N
                Me.Controls(r & i) = !Stationed
            Next
            
        ElseIf InStr(1, !Spot, "-") > 0 Then
        
        
        Else
            Me.Controls(Replace(!Spot, "-", "")) = !Stationed
        End If
    End If
    .MoveNext
Loop
End With

sExit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".onLoad"
    Resume Next
End Sub

