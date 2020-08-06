Attribute VB_Name = "Util"
Option Compare Database

Public Function createRelations()
    With CurrentDb
        Set rel = .CreateRelation(Name:="[rel.Name]", Table:="[rel.Table]", ForeignTable:="[rel.FireignTable]", Attributes:="[rel.Attributes]")
        rel.Fields.Append rel.CreateField("[fld.Name for relation]")
        rel.Fields("[fld.Name for relation]").ForeignName = "[fld.Name for relation]"
        .Relations.Append rel
    End With
End Function

Public Function saveRelations()

    For Each rel In CurrentDb.Relations
        With rel
            Debug.Print "Name: " & .Name
            Debug.Print "Attributes: " & .Attributes
            Debug.Print "Table: " & .Table
            Debug.Print "ForeignTable: " & .ForeignTable

            Debug.Print "Fields:"
            For Each fld In .Fields
                Debug.Print "Field: " & fld.Name
            Next
        End With
    Next
End Function

Public Function fixCase(ByRef s As Control)
On Error Resume Next
    s.Value = UCase(Left(s.Value, 1)) & Right(LCase(s.Value), Len(s.Value) - 1)
End Function

Public Function cETA(ByVal DOF, ATD, ETD, ETE As Date) As Date
    cETA = Format([DOF] + IIf([ATD] Is Null, [ETD], [ATD]) + [ETE], "Short Time")
End Function

Function LToZ(ByVal lcl As String) As Date
    Dim Timezone As Integer
    Timezone = DLookup("Timezone", "tblSettings")
    If DLookup("dst", "tblsettings") Then Timezone = Timezone + 1
    If lcl = "" Then Exit Function
    
    LToZ = DateAdd("h", -Timezone, lcl)
End Function

Function ZToL(ByVal zulu As String, Optional isTime As Boolean) As String
    Dim Timezone As Integer
    Timezone = DLookup("Timezone", "tblSettings")
    If DLookup("dst", "tblsettings") Then Timezone = Timezone + 1
    If zulu = "" Then Exit Function
    
    ZToL = DateAdd("h", Timezone, zulu)
    If isTime Then ZToL = Format(ZToL, "hh:nn")
End Function

Public Function isDST(ByVal d0 As Date, Optional locale As String = "US") As Boolean
Dim dstOn, dstOff As String

    Select Case locale
    
    Case "US"
        dstOn = "MAR 8 "
        dstOff = "NOV 1 "
        isDST = d0 >= NextSun("Mar 8 " & Year(d0)) And d0 < NextSun("Nov 1 " & Year(d0))
    Case "EU"
        dstOn = "MAR 8 "
        dstOff = "NOV 1 "
        isDST = d0 >= LastSun("Mar 8 " & Year(d0)) And d0 < LastSun("Nov 1 " & Year(d0))
        
    End Select

   
End Function

Private Function NextSun(D1 As Date) As Date
   NextSun = IIf(Weekday(D1) = 1, D1, D1 + 7 - (Weekday(D1) - 1))
End Function

Private Function LastSun(D1 As Date) As Date
    LastSun = D1 - IIf(Weekday(D1) = 1, 7, Weekday(D1) - 1)
End Function

Public Sub exportSchema(Optional Path As String = "%USERPROFILE%\Documents\GitHub\SCHEMA EXPORT\")
Dim f As String
Dim tdf As DAO.TableDef
Path = Replace(Path, "%USERPROFILE%", Replace(Environ$("userprofile"), "C:\", "D:\"))

    log "Creating path...", "Util.exportSchema"
    log IIf(Util.createPath(Path), "Success!", "Failed to create path."), "Util.exportSchema"
    
    Dim ad As additionalData: Set ad = Application.CreateAdditionalData
    For Each r In CurrentDb.Relations
    With r
        log "Gathering relationship: " & .Name, "Util.eportSchema"
        ad.add .Name
        ad.add .Attributes
        ad.add .Table
        ad.add .ForeignTable

        For Each fld In .Fields
            ad.add fld.Name
        Next
    End With: Next
    
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.Name, 3) = "tbl" Then
            log "Exporting schema: " & tdf.Name, "Util.exportSchema"
            Application.ExportXML acExportTable, tdf.Name, , Path & tdf.Name & ".xsd", , , , , , ad
            
            DoEvents
        End If
    Next
    log "Done!", "Util.exportSchema"
    
End Sub

Public Sub exportAllCode()
Dim c As VBComponent
Dim Sfx, exportLocation As String
Dim num As Integer

'If dir(exportLocation) = "" Then createPath exportLocation

    For Each c In Application.VBE.VBProjects(1).VBComponents
        exportLocation = CurrentProject.Path & "\DB EXPORT\"

        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            log "Exporting " & c.Name & Sfx & "...", "Util.ExportAllCode"
            DoEvents

            Select Case Left(c.Name, InStr(1, c.Name, "_"))
                Case "Form_"
                    exportLocation = exportLocation & "Forms\"
                Case "Report_"
                    exportLocation = exportLocation & "Reports\"
                Case Else
                    exportLocation = exportLocation & "Modules\"
            End Select

            If dir(exportLocation) = "" Then createPath exportLocation
            c.Export exportLocation & "\" & c.Name & Sfx
            num = num + 1
        End If
    Next c

    log "Done! Successfully exported " & num & " objects.", "Util.exportAllCode"
End Sub


Public Sub exportAllObjects()
On Error GoTo errtrap
Dim db As DAO.Database
Dim td As TableDef
Dim d As Document
Dim c As Container
Dim i As Integer
Dim exportLocation, pDir, subDir As String

Set db = CurrentDb
pDir = CurrentProject.Path & "\DB EXPORT\"
'If dir(exportLocation) = "" Then createPath exportLocation

    log "Exporting Tables...", "Util.exportAllObjects"
    
        subDir = "table\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        For Each td In db.TableDefs
            If Left(td.Name, 3) = "tbl" Then
                DoCmd.TransferText acExportDelim, , td.Name, exportLocation & td.Name & ".txt", True
            End If
        Next
    
    log "Exporting Forms...", "Util.exportAllObjects"
    
        subDir = "form\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        Set c = db.Containers("Forms")
        For Each d In c.Documents
            Application.SaveAsText acForm, d.Name, exportLocation & "Form_" & d.Name & ".txt"
        Next d
    
    log "Exporting Reports...", "Util.exportAllObjects"
    
        subDir = "report\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        Set c = db.Containers("Reports")
        For Each d In c.Documents
            Application.SaveAsText acReport, d.Name, exportLocation & "Report_" & d.Name & ".txt"
        Next d
    
    log "Exporting Scripts...", "Util.exportAllObjects"
    
        subDir = "scripts\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        Set c = db.Containers("Scripts")
        For Each d In c.Documents
            Application.SaveAsText acMacro, d.Name, exportLocation & "Macro_" & d.Name & ".txt"
        Next d
    
    log "Exporting Modules...", "Util.exportAllObjects"
    
        subDir = "module\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        Set c = db.Containers("Modules")
        For Each d In c.Documents
            Application.SaveAsText acModule, d.Name, exportLocation & "Module_" & d.Name & ".txt"
        Next d
    
    log "Exporting Queries...", "Util.exportAllObjects"
    
        subDir = "query\"
        exportLocation = pDir & subDir
        If dir(exportLocation) = "" Then createPath exportLocation
        
        For i = 0 To db.QueryDefs.Count - 1
            Application.SaveAsText acQuery, db.QueryDefs(i).Name, exportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
        Next i
    
    Set db = Nothing
    Set c = Nothing
    
    log "All database objects have been exported as a text file to " & exportLocation, "Util.exportAllObjects"

sExit:
    Exit Sub
errtrap:
    Util.ErrHandler err, Error$, "Util.EXPORT"
    Resume sExit
End Sub

Public Function findRank(ByVal r As String) As Variant
For i = 0 To 15
    If rankID(i) = r Then
        findRank = i
        Exit Function
    End If
Next
findRank = Null
End Function

Public Function rankID(ByVal i As Integer, Optional ByVal tree As Integer = 0) As String
Select Case tree
    Case 0
        rankID = Array("AB", "Amn", "A1C", "SrA", "SSgt", "TSgt", "MSgt", "SMSgt", "CMSgt", "2dLt", "1Lt", "Capt", "Maj", "LtCol", "Col")(i)
    
    Case 1
        rankID = "GS-" & Format(i, "00")
        
End Select
End Function

Public Function getOpInitials(Optional ByVal username As String) As String
    getOpInitials = Nz(DLookup("opInitials", "tbluserauth", "username = '" & IIf(username <> "", username, Environ$("username")) & "'"))
End Function

Public Function getUSN(Optional ByVal opInitials As String) As String
    If opInitials <> "" Then
        getUSN = DLookup("username", "tblUserAuth", "opInitials ='" & opInitials & "'")
    Else
        getUSN = Environ$("username")
    End If
End Function

Public Function createPath(ByVal Path As String) As Boolean
On Error GoTo errtrap
If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)

Dim strSlash, strFolder, strRSFolder As String
Dim fs, cf, X
Set fs = CreateObject("Scripting.FileSystemObject")
strSlash = "\"
intCurrPos = 4
strFolder = Path
intLength = Len(strFolder)

If Left(Path, 2) = "\\" Then intCurrPos = InStr(InStr(3, Path, strSlash) + 1, Path, strSlash)
    
    If intLength > 3 Then
        Do
            intNextPos = InStr(intCurrPos, strFolder, strSlash)
            intCurrPos = intNextPos + 1
            
            If intNextPos > 0 Then
            
               If fs.FolderExists(Left(strFolder, intNextPos - 1)) = False Then
                  Set cf = fs.CreateFolder(Left(strFolder, intNextPos - 1))
               End If
            Else
            
               If fs.FolderExists(Left(strFolder, intLength)) = False Then
                  Set cf = fs.CreateFolder(Left(strFolder, intLength))
               End If
            End If
        Loop Until (intNextPos = 0)
    End If
    
fExit:
    createPath = True
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.createPath"
End Function

Public Function testBackend(ByVal key As String, ByVal backend As String) As Boolean
On Error GoTo errtrap
Dim RS As DAO.Recordset
Dim db As DAO.Database
Set db = CurrentDb

    Set RS = db.OpenRecordset("SELECT key FROM tblSettings IN '" & backend & "'", , dbFailOnError)
    With RS
        If Not .Fields("key").Name = "key" Then
            MsgBox "Invalid AeroStat Backend format. (Key mismatch)", vbCritical, "AeroStat"
            log "Invalid AeroStat Backend file. (Key mismatch)", "Util.testBackend", "WARN"
            Exit Function
'        ElseIf .EOF Then
'            MsgBox "Invalid AeroStat Backend file.", vbCritical, "AeroStat"
'            log "Invalid AeroStat Backend file.", "Util.testBackend", "FATAL"
'            Exit Function
        End If
        .Close
    End With
    
    testBackend = Not Nz(backend) = ""
fExit:
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.testBackend"
End Function

Public Function relinkTables(Optional ByVal backend As String, Optional ByRef loading As Variant = Null) As Boolean
On Error GoTo errtrap
Dim db As DAO.Database
Dim RS As DAO.Recordset
Dim tdf As DAO.TableDef
Dim ld As Boolean
Dim strTable, lnkDatabase, newFile, key As String
Set db = CurrentDb

If IsNull(loading) Then
    DoCmd.OpenForm "frmLoading"
    Set loading = Forms!frmLoading
    loading!loadingText.Caption = "Validating data..."
    DoEvents
    ld = True
End If
DoEvents
'lnkDatabase = Nz(backend, _
'                Nz(DLookup("backend", "lclver"), _
'                    Nz(DLookup("backend", "tblSettings")) _
'                ) _
'            )
If IsMissing(backend) Or Nz(backend) = "" Then
    lnkDatabase = Nz(DLookup("backend", "lclver"), _
                        Nz(DLookup("backend", "tblSettings")) _
                    )
Else
    lnkDatabase = backend
End If
key = DLookup("key", "lclver")

    'If dir(lnkDatabase) = "" Or dir(lnkDatabase, vbDirectory) = "." Then
    If Not testBackend(key, lnkDatabase) Then
        GoTo showErr
show:
        Dim fd As Office.FileDialog
        Set fd = Access.FileDialog(msoFileDialogOpen)
        With fd
            .title = "Select BACKEND file"
            .Filters.clear
            .Filters.add "All Files", "*.*"

            On Error GoTo showErr
            If .show Then
'                Dim objAccess As Object
'                Dim objRecordset As Object
'                Set objAccess = New Access.Application
                
                For Each varfile In .SelectedItems
                    lnkDatabase = varfile
                Next
                If Not testBackend(key, lnkDatabase) Then GoTo showErr
                
'                db.Execute "UPDATE settings, lclver SET settings.backend = '" & lnkDatabase & "', lclver.backend = '" & lnkDatabase & "'"
'                objAccess.OpenCurrentDatabase lnkDatabase
'                Set objRecordset = objAccess.CurrentProject.Connection.Execute("tblSettings")
'                If objRecordset.Fields("key") <> key Then
'                    MsgBox "Invalid AeroStat Backend format. (Key mismatch)", vbCritical, "AeroStat"
'                    GoTo show
'                End If
            Else
                log "Cancelled by user.", "Util.relinkTables"
                GoTo fExit
            End If
        End With
    Else
        'GoTo funcExit
    End If
    
    
On Error GoTo errtrap
        If Not IsNull(loading) Then
            With loading!pBar
                loading!loadingText.Caption = "Updating tables..."
                DoEvents
                .Max = db.TableDefs.Count
                .Value = 0
            End With
        End If
        
        For Each tdf In db.TableDefs
            If Len(tdf.Connect) > 1 Then 'Only relink linked tables
                If tdf.Connect <> ";DATABASE=" & lnkDatabase Then 'only relink tables if they are not linked correctly
                    If Left(tdf.Connect, 4) <> "ODBC" And Left(tdf.Connect, 3) <> "WSS" Then 'Don't want to relink any ODBC tables
                        strTable = tdf.Name
                        'db.TableDefs(strTable).Connect = "MS Access;PWD=" & DBPassword & ";DATABASE=" & LnkDataBase
                        db.TableDefs(strTable).Connect = ";DATABASE=" & lnkDatabase
                        db.TableDefs(strTable).RefreshLink
                        'log tdf.Name & " refreshed.", "Util.relinkTables"
                    End If
                End If
'                If Nz(DLookup("sharepoint", "tblSettings"), False) And Left(tdf.Connect, 3) = "WSS" Then
'
'                    db.TableDefs(tdf.Name).RefreshLink
'                    log "SP Table " & tdf.Name & " refreshed.", "Util.relinkTables"
'                End If
            End If
            
            If Not IsNull(loading) Then
                With loading!pBar
                    .Value = .Value + 1
                End With
            End If
            DoEvents
        Next tdf
        
    '    Set rs = db.OpenRecordset("lclver")
    '    With rs
    '        .edit
    '        !backend = lnkDatabase
    '        .update
    '    End With
    
    DoEvents
    log "Table links re-synced.", "Util.relinkTables"
    db.Execute "UPDATE tblSettings SET backend = '" & lnkDatabase & "'"
    db.Execute "UPDATE lclver SET backend = '" & lnkDatabase & "'"
    relinkTables = True
    
fExit:
    If ld Then DoCmd.Close acForm, "frmLoading"
    Exit Function
errtrap:
    If err = 52 Then Resume Next
    ErrHandler err, Error$, "Util.relinkTables"
    Resume Next
showErr:
    ErrHandler err, Error$, "Util.relinkTables"
    MsgBox "Invalid AeroStat Backend format.", vbCritical, "AeroStat"
    GoTo show
End Function

Public Function serialDate(d As Date) As Date
    serialDate = DateSerial(Year(d), Month(d), Day(d))
End Function

Public Function noBreaks(ByVal s As String, Optional ByVal separator As String) As String
    s = Trim(s)
    noBreaks = Replace(s, vbCrLf, Nz(separator, " "))
End Function

Public Function break() As String
break = vbCrLf
End Function

Public Sub ErrHandler(err As Integer, msg As String, Optional frm As String)
    Debug.Print Format(Now, "dd-mmm-yy hh:nnL ") & "(" & err & ") " & IIf(IsNull(frm), "", "[" & frm & "] ") & msg
    log "(" & err & ") " & msg, IIf(IsNull(frm), "", "[" & frm & "] "), "WARN"
End Sub

Public Sub log(msg As String, module As String, Optional priority As String = "INFO")
On Error Resume Next
Dim db As DAO.Database
Set db = CurrentDb
Dim sql As String
sql = "INSERT INTO debug (username,initials,computername,priority,module,details) SELECT tblUserAuth.username, tblUserAuth.opInitials, tblUserAuth.lastsystem, '" & priority & "', '" & module & "', " & """" & msg & """" & " FROM tblUserAuth WHERE tblUserAuth.username = '" & getUSN & "';"

    db.Execute sql
    Debug.Print Format(Now, "dd-mmm-yy hh:nn:ssL") & "[" & priority & "] " & module & ": " & msg
End Sub

Public Function base7(ByVal num As Integer) As Integer
    base7 = num
    If num < 0 Then base7 = num + 7
End Function

Public Function onShareDriveDisconnect() As Boolean
'Eventually, this will control pending changes to be made once the drive reconnects.
onShareDriveDisconnect = (MsgBox("The Shared Drive was disconencted. Retry?", vbQuestion + vbYesNo, "AeroStat") = vbYes)
End Function

Public Function convertFt(ByVal f As Variant, ByVal toDecimal As Boolean) As Variant
    If toDecimal Then
        convertFt = Left(f, InStr(1, f, "’") - 1) + Mid(f, InStr(1, f, "’") + 1, 2) / 12
    End If
End Function

Public Function setting(ByVal v As String) As Variant
    setting = DLookup(v, "tblSettings")
End Function

Public Function appendUser(ByVal usr As String, ByVal fld As String, ByVal v As Variant)
On Error GoTo errtrap
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & usr & "';")
With RS
    If .EOF Then Exit Function
    .edit
    .Fields(fld) = v
    .Update
    .Close
End With
Set RS = Nothing

sExit:
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.appendUser"

End Function

Public Function syncTrafficLog(Optional ByVal recID As Integer, Optional ByVal tbl As String, Optional ByVal newrec As Boolean, Optional M As String)
On Error Resume Next
Dim RS As DAO.Recordset
Dim rsAlert As DAO.Recordset
t = Now

If Not IsNull(newrec) Then
Set rsAlert = CurrentDb.OpenRecordset("tblTrafficLogAlert")
Set rstbl = CurrentDb.OpenRecordset("SELECT * FROM " & tbl & " WHERE ID = " & recID)
    With rsAlert
        .AddNew
        !timestamp = t
        !PID = recID
        !opInitials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
        With rstbl: Select Case tbl
            Case "Traffic"
                rsAlert!alerttype = 1
                rsAlert!msg = IIf(newrec, "[+]: ", IIf(!Status = "Cancelled", "[-]: ", "[*]: ")) & !Callsign & "/" & !Type & " | " & !Status
                
            Case "tblPPR"
                rsAlert!alerttype = 2
                rsAlert!msg = "[PPR] " & !Callsign & "/" & !Type & " | " & Format(!arrDate, "dd/hh:nn") & "L-" & Format(!arrDate, "dd/hh:nn") & "L | " & !Status
                
            Case "Custom"
                rsAlert!alerttype = 3
                rsAlert!msg = M
            End Select
        End With
        .Update
        .Bookmark = .LastModified
    End With
    
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
'    With rs
'        .edit
'        !frmtrafficlogalert = rsAlert!ID
'        .update
'    End With
End If

Set RS = CurrentDb.OpenRecordset("tblSettings")

    RS.edit
    RS!frmTrafficLogSync = t
    RS.Update
    Set RS = CurrentDb.OpenRecordset("lclver")
    RS.edit
    RS!frmTrafficLogSync = t
    RS.Update
    RS.Close

Set RS = Nothing
End Function

Public Function findParentByTail(ByVal Tail As String, ByVal dir As Integer) As Integer
'dir = 1 - in, 2 - out, 3 - local
findParentByTail = 0

Dim RS As DAO.Recordset
    Select Case dir
        Case 1
            Set RS = CurrentDb.OpenRecordset("qInbound")
        Case 2
            Set RS = CurrentDb.OpenRecordset("qOutbound")
        Case 3
            Set RS = CurrentDb.OpenRecordset("qLocal")
    End Select
    
    With RS
    Do While Not .EOF
        If !Tail = Tail Then
            findParentByTail = !ID
            Exit Function
        End If
        .MoveNext
    Loop
    .Close
    End With
    Set RS = Nothing
End Function

Public Function isShiftClosed(ByVal user As String) As Integer
isShiftClosed = 0
Dim result As Boolean
result = Nz(DLookup("closed", "tblshiftmanager", "shiftID = " & Nz(DLookup("lastShift", "tbluserauth", "username = '" & user & "'"), 0)), True)
If Not result Then isShiftClosed = DLookup("lastShift", "tbluserauth", "username = '" & user & "'")
End Function

Public Function getLogName(ByVal oi As String) As String
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE opInitials = '" & oi & "';")
If RS.RecordCount = 0 Then
    getLogName = ""
    Exit Function
End If

With RS
    getLogName = Left(!firstName, 1) & ". " & !lastName & "/" & !opInitials
    .Close
End With
End Function

Public Function fullName(ByVal oi As String) As String
fullName = DLookup("rankID", "tblUserAuth", "opInitials = '" & oi & "'") & " " & DLookup("firstName", "tblUserAuth", "opInitials = '" & oi & "'") & " " & DLookup("lastName", "tblUserAuth", "opInitials = '" & oi & "'")
End Function

Public Function getPos(ByVal perm As Integer, Optional ByVal alt As Boolean = False) As String
    Select Case perm
        Case 0, 99
            getPos = ""
        Case 1
            getPos = "(AFM)"
        Case 2
            getPos = "(DAFM)"
            If alt Then getPos = "(AAFM)"
        Case 3
            getPos = "(NAMO)"
            If alt Then getPos = "(AMOM)"
        Case 4
            getPos = "(NAMT)"
            If alt Then getPos = "(AMOT)"
        Case 5
            getPos = "(AMOS)"
        Case 6
            getPos = "(AMSL)"
        Case 7
            getPos = "(AMOC)"
        Case 8
            getPos = "(TRAINEE)"
    End Select
End Function

Public Function getDuty(ByVal perm As Integer, Optional ByVal alt As Boolean = False) As String
    Select Case perm
        Case 0, 99
            getDuty = ""
        Case 1
            getDuty = "Airfield Manager"
        Case 2
            getDuty = "Deputy Airfield Manager"
            If alt Then getDuty = "Assistant Airfield Manager"
        Case 3
            getDuty = "NCOIC Airfield Management Operations"
            If alt Then getDuty = "Airfield Management Operations Manager"
        Case 4
            getDuty = "NCOIC Airfield Management Training"
            If alt Then getDuty = "Airfield Management Training Manager"
        Case 5
            getDuty = "Airfield Management Operations Supervisor"
        Case 6
            getDuty = "Airfield Management Shift Lead"
        Case 7, 8
            getDuty = "Airfield Management Operations Coordinator"
    End Select
End Function

Public Function getAccessSP() As Boolean
On Error GoTo getAccessSP_err
Dim RS As DAO.Recordset
Dim qdf As QueryDef
getAccessSP = True

    Set qdf = CurrentDb.QueryDefs("qMissionTracker")
    qdf.Parameters("fromDay") = LToZ(Date)
    Set RS = qdf.OpenRecordset()
    
    Set RS = Nothing

getAccessSp_Exit:
    Exit Function

getAccessSP_err:
    If err = 3011 Or err = 3841 Then
        MsgBox "SharePoint connection was rejected." & vbCrLf & "Changes to flight plans will not be sent to the SharePoint." & vbCrLf & vbCrLf & "Error code: " & err, vbCritical
        getAccessSP = False
        Resume Next
    Else
        ErrHandler err, Error$, "Util.getAccessSP"
        getAccessSP = False
        Exit Function
    End If
End Function

Public Function cnlFlight(ByVal rid As Recordset) As Boolean
Dim RS As DAO.Recordset
Dim rstSP As DAO.Recordset
Set RS = rid
bClose = True
With RS

    If Not .EOF Then
        .edit
        Select Case !Status
            Case "Closed"
                MsgBox "This flight plan is already closed.", vbInformation, "Error"
                Exit Function
                
            Case "Cancelled"
                If MsgBox("Re-activate flight?", vbYesNo + vbQuestion, "AeroStat") = vbYes Then
                    !Status = "Pending"
                    cnlFlight = True
                End If
            
            Case "Pending", "-"
                If MsgBox("Cancel flight?", vbYesNo + vbQuestion, "AeroStat") = vbYes Then
                    !Status = "Cancelled"
                    cnlFlight = True
                End If
                
            Case "Enroute"
                If MsgBox("Cancel flight?" & vbCrLf & "(This will reset any previously entered times)", vbYesNo + vbQuestion, "AeroStat") = vbYes Then
                    !Status = "Cancelled"
                    cnlFlight = True
                End If
        End Select
        .Update
        syncTrafficLog !ID, "Traffic", False
    Else
        MsgBox "Flight not found", vbInformation, "Error"
    End If
End With

End Function

'Public Function checkParking(ByVal Callsign As String, ByVal Tail As String) As Integer
'Dim rs As DAO.Recordset
'Dim stn As String
'stn = DLookup("Station", "tblSettings")
'Set rs = CurrentDb.OpenRecordset("qOnStation")
'
'    Do While Not rs.EOF
'        If rs!Tail = Tail Or rs!Callsign = Callsign Then
'            rs.edit
'            checkParking = rs!ID
'            rs!Stationed = False
'            rs.update
'            Exit Function
'        End If
'        rs.MoveNext
'    Loop
'
'End Function

Public Function numTense(ByVal num As String) As String
Dim tns As String
If Not IsNumeric(num) Then numTense = num: Exit Function

    If Val(Right(num, 2)) >= 11 And Val(Right(num, 2)) <= 13 Then
        tns = "th"
    Else: Select Case Right(Val(num), 1)
            Case 0, 4 To 9
                tns = "th"
            Case 1
                tns = "st"
            Case 2
                tns = "nd"
            Case 3
                tns = "rd"
        End Select
    End If
    
    numTense = num & tns
End Function

Public Function isT(ByVal t As String, ByVal acType As String) As Boolean
On Error Resume Next
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("tblBaseAcft")
isT = True

If Not Left(t, 1) = "N" And Not IsNumeric(Right(t, 1)) Then t = Left(t, Len(t) - 1)

With RS
    Do While Not .EOF
        If Right(t, 4) = Right(!Tail, 4) And !acType = acType Then
            isT = False
            Exit Do
        End If
        .MoveNext
    Loop
    .Close
End With

Set RS = Nothing
        
End Function

Public Function test(str As String) As String
test = Left(str, InStr(1, str, ",") - 1)

End Function
