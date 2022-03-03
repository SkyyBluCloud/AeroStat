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

Private Sub Form_Close()
    Call unloadSettings
End Sub

Private Sub Form_Open(cancel As Integer)
    Call loadSettings
End Sub

Private Sub unloadSettings()
On Error GoTo errtrap
    Set Me.Recordset = Nothing
    CurrentDb.TableDefs.delete "tempSettings"
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".Form_Close"
End Sub

Private Sub loadSettings()
On Error GoTo errtrap
Dim db As DAO.Database: Set db = CurrentDb
Dim tempTbl As TableDef
Dim rsSettings As DAO.Recordset: Set rsSettings = db.OpenRecordset("tblSettings")

    'Create a temptbl; load the keys from settings as temptbl fields
    Set tempTbl = db.CreateTableDef("tempSettings")
    
    With rsSettings: Do While Not .EOF
        tempTbl.Fields.Append tempTbl.CreateField(!key, .Fields("datatype"))
        .MoveNext
    Loop
    End With
    
    db.TableDefs.Append tempTbl
    
    'Load the temptbl recordset
    Dim rsTT As DAO.Recordset: Set rsTT = db.OpenRecordset("tempSettings")
    With rsTT
        'Go to the start of settings
        rsSettings.MoveFirst
        .AddNew 'Add new (single) record
        
        'Load setting values to the temptbl fields
        Do While Not rsSettings.EOF
            .Fields(rsSettings!key).Value = rsSettings!data
            rsSettings.MoveNext
        Loop
        
        .update
     End With
     
    'Set the forms recordset to the completed recordset
    Set Me.Recordset = rsTT.OpenRecordset
    
    rsSettings.Close
    rsTT.Close
sexit:
    Set rsSettings = Nothing
    Set rsTT = Nothing
    Set tempTbl = Nothing
    Exit Sub
    
errtrap:
    ErrHandler err, Error$, Me.Name & ".reloadSettings"
    GoTo sexit
End Sub
