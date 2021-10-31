Attribute VB_Name = "TrafficUtil"
Option Compare Database
Option Explicit

Public Function getArrDate(ByVal DOF As Date, ByVal ATD As Date, _
                            ByVal ETD As Date, ByVal ETE As Date, _
                            ByVal cETA As Date) As Date
Dim tz As Integer: tz = DLookup("data", "tblSettings", "key = ""timezone""")

    getArrDate = Format(DateAdd("h", tz, (DOF + (Nz(ATD, ETD) + ETE))), "dd-mmm-yy") & " " & Format(DateAdd("h", tz, cETA), "hh:nn")
    
End Function

Public Function atlasPull(ByVal varDate As Variant) As Boolean
On Error GoTo errtrap
log "Start Atlas pull for " & varDate, "TrafficUtil.atlasPull"
Dim db As DAO.Database: Set db = CurrentDb
Dim rsAtlas As DAO.Recordset, rsTraffic As DAO.Recordset
Dim qdfAtlas As DAO.QueryDef: Set qdfAtlas = db.QueryDefs("qAtlasPull")

If Not IsDate(varDate) Then varDate = Date
qdfAtlas.Parameters("varDate") = varDate
Set rsAtlas = qdfAtlas.OpenRecordset
Set rsTraffic = db.OpenRecordset("tblTraffic")
    
With rsTraffic

    
    Do While Not rsAtlas.EOF
    
        If Right(rsAtlas!depPoint, 3) = Right(DLookup("data", "tblsettings", "key = 'station'"), 3) Or Right(rsAtlas!Destination, 3) = Right(DLookup("data", "tblsettings", "key = 'station'"), 3) Then
    
            If IsNull(DLookup("atlasID", "tblTraffic", "atlasID = " & rsAtlas!atlasID)) Then 'Atlas record does not exist locally; create it
                .AddNew
                
            Else 'Atlas record exists locally; update it
            
                .FindFirst "atlasID = " & rsAtlas!atlasID
                If Not .EOF Then
                    .edit
                Else 'Something weird happened
                    GoTo fexit
                End If
            End If
            
            With rsAtlas
    
                Dim fld: For Each fld In .Fields
                    If Left(fld.Name, 2) <> "In" And Left(fld.Name, 2) <> "Ot" Then
                    
                        Select Case True
                        
'                            Case fld.Name = "ETA"
'                                If Not IsDate(fld.Value) Then
'                                    rsTraffic.Fields(fld.Name).Value = cETA(!DOF, !ETD, !ETE, , !ATD, !ATA)
'                                Else
'                                    rsTraffic.Fields(fld.Name).Value = Nz(fld.Value)
'                                End If
'                                If TimeValue(rsTraffic.Fields(fld.Name).Value) = #12:00:00 AM# Then rsTraffic.Fields(fld.Name).Value = Null
                                
    '                        Case fld.Name = "ATA"
    '                            If fld.Value <> "" Then rsTraffic.Fields(fld.Name).Value = Nz(fld.Value)
    '
'                            Case fld.Name = "ETD"
'                                If Not IsDate(fld.Value) Then
'                                    rsTraffic.Fields(fld.Name).Value = CDate(TimeValue(!ETA) - TimeValue(!ETE))
'                                Else
'                                    rsTraffic.Fields(fld.Name).Value = Nz(fld.Value)
'                                End If
    '
    '                        Case fld.Name = "ATD"
    '                            If Nz(fld.Value) = "" Then
    '                                rsTraffic.Fields(fld.Name) = Null
    '                            Else
    '                                rsTraffic.Fields(fld.Name).Value = fld.Value
                                'End If
                                
                            Case fld.Name = "STS"
                                rsTraffic!Status = fld.Value
                            
                            Case fld.Name = "Destination", fld.Name = "depPoint"
                                rsTraffic.Fields(fld.Name).Value = IIf(Len(Nz(fld.Value)) = 3, "K" & fld.Value, Nz(fld.Value))
                            
                            Case Not fld.Name Like "Expr*" And fld.Name <> "LastUpdated"
                                rsTraffic.Fields(fld.Name).Value = Nz(fld.Value)
                            
                        End Select
                        
    '                    If Nz(rsTraffic!Tail) = "" Then
    '                        rsTraffic!Tail = "ats" & Right(rsAtlas!atlasID, 3)
    '                    End If
                        
                    End If
                Next fld
            End With
            'log Nz(rsTraffic!Tail, rsTraffic!Callsign), "TrafficUtil.atlasPull"
            .update
        End If
        
        rsAtlas.MoveNext
    Loop
    
End With

fexit:
    log "Success!", "TrafficUtil.atlasPull"
    atlasPull = True
    Exit Function
    Resume Next
errtrap:
    'log fld.Name & " " & fld.Value, "TrafficUtil.atlasPull", "ERR"
    ErrHandler err, Error$, "TrafficUtil.atlasPull"
End Function

Public Function linkAtlas(ByVal newrec As Boolean, ByVal atlasID As Double) As Double
On Error GoTo errtrap
Dim rsConv As DAO.Recordset: Set rsConv = CurrentDb.OpenRecordset("tblAtlasConversion")
Dim rsAtlas As DAO.Recordset: Set rsAtlas = CurrentDb.OpenRecordset("atlAtlas")
    
    'Create a new Atlas record, or find the existing one
    With rsAtlas
    If newrec Or atlasID = 0 Then
        .AddNew
    Else
        .FindFirst "recID = " & atlasID
        If Not .EOF Then
            .edit
        Else
            GoTo fexit
        End If
    End If
    End With
    
    If Not CurrentProject.AllForms("quick_input").IsLoaded Then DoCmd.OpenForm "quick_input", , , "atlasID = " & atlasID, acFormEdit, acHidden
    
    'Go through the atlas fields and append the solution for each
    With rsConv: Do While Not .EOF
        rsAtlas.Fields(!atlasfield).Value = Eval(!atlassolution)
        .MoveNext
    Loop: End With
    
    Forms!quick_input.Visible = False
    
    'Update and link
    With rsAtlas
        .update
        .Bookmark = .LastModified
        'atlasID = !recID
        linkAtlas = !recID
        .Close
    End With
    
fexit:
    If linkAtlas <> 0 Then log "Update ATLAS Link! (" & linkAtlas & ")", "TrafficUtil.linkAtlas"
    Set rsConv = Nothing
    Set rsAtlas = Nothing
    Exit Function
    
errtrap:
    ErrHandler err, Error$, "TrafficUtil.linkAtlas"
    GoTo fexit
    Resume Next
    
End Function
