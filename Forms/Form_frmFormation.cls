VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFormation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub addForm_Click()
Dim es As Boolean
Dim rss As DAO.Recordset
Dim rsi As DAO.Recordset
Dim fld As DAO.Field

    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            If IsNull(ctl) Then
            
                ctl.BorderColor = RGB(255, 0, 0)
                es = True
            Else
                ctl.BorderColor = RGB(180, 180, 180)
            End If
        End If
    Next
    If es Then Exit Sub
    
    Set rsi = Me.RecordsetClone
    Set rss = rsi.Clone
    
    With rss
        If .RecordCount > 0 Then
            .Bookmark = Me.Bookmark
            With rsi
                .AddNew
                For Each fld In rss.Fields
                    With fld
                        If .Attributes And dbAutoIncrField Then
                            'Do nothing
                        ElseIf .Name = "Callsign" Or .Name = "Type" Or .Name = "Tail" Then
                            'Do nothing
                        ElseIf .Name = "FormID" Then
                            rsi.Fields(.Name) = FormID
                        Else
                            rsi.Fields(.Name) = .Value
                        End If
                    End With
                Next
                .update
                .Close
            End With
        End If
        .Close
    End With
    
    Set rsi = Nothing
    Set rss = Nothing
    formList.Requery
    
End Sub

Private Sub Form_Load()
formList.RowSource = "SELECT Traffic.Callsign, Traffic.Type, Traffic.Tail FROM Traffic WHERE (((Traffic.FormID)=" & ID & "));"
header.Caption = Replace(header.Caption, "<CALL>", Callsign)
End Sub


