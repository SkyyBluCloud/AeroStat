Attribute VB_Name = "TrafficUtil"
Option Compare Database
Option Explicit

Public Function getArrDate(ByVal DOF As Date, ByVal ATD As Date, _
                            ByVal ETD As Date, ByVal ETE As Date, _
                            ByVal cETA As Date) As Date
Dim tz As Integer: tz = DLookup("data", "tblSettings", "key = ""timezone""")

    getArrDate = Format(DateAdd("h", tz, (DOF + (Nz(ATD, ETD) + ETE))), "dd-mmm-yy") & " " & Format(DateAdd("h", tz, cETA), "hh:nn")
    
End Function

Public Function atlasPull(Optional ByVal varDate As Variant) As Boolean
On Error GoTo errtrap
Dim rsAtlas As DAO.Recordset: Set rsAtlas = CurrentDb.OpenRecordset("atlAtlas")
If Not IsDate(varDate) Then varDate = Date



fexit:
    Exit Function
    
errtrap:
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
    
    'Go through the atlas fields and append the solution for each
    With rsConv: Do While Not .EOF
        rsAtlas.Fields(!atlasfield).Value = Eval(!atlasSolution)
        .MoveNext
    Loop: End With
    
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
