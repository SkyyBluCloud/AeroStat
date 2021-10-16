VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSettingsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(cancel As Integer)
    Call reloadSettings
    
End Sub

Private Sub reloadSettings()
On Error GoTo errtrap
Dim tempTbl As New TableDef
Dim rsSettings As DAO.Recordset: Set rsSettings = CurrentDb.OpenRecordset("tblSettings")

    'Create a temptbl; load the keys from settings as temptbl fields
    With rsSettings: Do While Not .EOF
        tempTbl.Fields.Append tempTbl.CreateField(!key)
        .MoveNext
    Loop

    End With
    
    'Load the temptbl recordset
    Dim rsTT As DAO.Recordset: Set rsTT = tempTbl.OpenRecordset
    With rsTT
        'Go to the start of settings
        rsSettings.MoveFirst
        .AddNew 'Add new (single) record
        
        'Load setting values to the temptbl fields
        Do While Not rsSettings.EOF
            rsTT.Fields(!key).Value = !Data
            .MoveNext
          Loop
     End With
     
    'Set the forms recordset to the completed recordset
    Set Me.Recordset = rsTT.OpenRecordset
    
sexit:
    rsSettings.Close
    rsTT.Close
    Set rsSettings = Nothing
    Set rsTT = Nothing
    Set tempTbl = Nothing
    Exit Sub
    
errtrap:
    ErrHandler err, Error$, Me.Name & ".reloadSettings"
End Sub
