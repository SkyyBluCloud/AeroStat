Attribute VB_Name = "Util"
Option Compare Database

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
On Error GoTo errTrap
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

sexit:
    Exit Sub
errTrap:
    Util.errHandler err, Error$, "Util.EXPORT"
    Resume sexit
End Sub

Public Function findRank(ByVal r As String) As Variant
For i = 0 To 15
    If rank(i) = r Then
        findRank = i
        Exit Function
    End If
Next
findRank = Null
End Function

Public Function rank(ByVal i As Integer) As String
rank = Array("Civ", "AB", "Amn", "A1C", "SrA", "SSgt", "TSgt", "MSgt", "SMSgt", "CMSgt", "2dLt", "1Lt", "Capt", "Maj", "LtCol", "Col")(i)
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
On Error GoTo errTrap
If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)

Dim strSlash, strFolder, strRSFolder As String
Dim fs, cf, x
Set fs = CreateObject("Scripting.FileSystemObject")
strSlash = "\"
intCurrPos = 4
strFolder = Path
intLength = Len(strFolder)
    
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
errTrap:
    errHandler err, Error$, "Util.createPath"
End Function

Private Function testBackend(ByVal key As String, ByVal backend As String) As Boolean
On Error GoTo fExit
Dim rs As DAO.Recordset
Dim db As DAO.Database
Set db = CurrentDb

    Set rs = db.OpenRecordset("SELECT key FROM settings IN '" & backend & "'", , dbFailOnError)
    With rs
        If .EOF Then
            MsgBox "Invalid AeroStat Backend file.", vbCritical, "AeroStat"
            Exit Function
        ElseIf !key <> key Then
            MsgBox "Invalid AeroStat Backend format. (Key mismatch)", vbCritical, "AeroStat"
            Exit Function
        End If
        .Close
    End With
    
    testBackend = True
fExit:
errHandler err, Error$, "Util.testBackend"
End Function

Public Function relinkTables() As Boolean
On Error GoTo errTrap
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
Dim tdf As DAO.TableDef
Dim strTable, lnkDatabase, newFile, key As String
lnkDatabase = DLookup("backend", "lclver")
key = DLookup("key", "lclver")
Set dbs = CurrentDb()
    
    'If dir(lnkDatabase) = "" Or dir(lnkDatabase, vbDirectory) = "." Then
    If Not testBackend(key, Nz(lnkDatabase)) Then
show:
        Dim fd As Office.FileDialog
        Set fd = Access.FileDialog(msoFileDialogFilePicker)
        With fd
            .title = "Please select the backend file."
            .Filters.clear
            .Filters.add "Access Databases", "*.accdb"

On Error GoTo showErr
            If .show Then
'                Dim objAccess As Object
'                Dim objRecordset As Object
'                Set objAccess = New Access.Application
                
                For Each varfile In .SelectedItems
                    lnkDatabase = varfile
                Next
                If Not testBackend(key, Nz(lnkDatabase)) Then GoTo show
                
'                objAccess.OpenCurrentDatabase lnkDatabase
'                Set objRecordset = objAccess.CurrentProject.Connection.Execute("settings")
'                If objRecordset.Fields("key") <> key Then
'                    MsgBox "Invalid AeroStat Backend format. (Key mismatch)", vbCritical, "AeroStat"
'                    GoTo show
'                End If
            Else
                log "Cancelled by user.", "Util.relinkTables"
                Exit Function
            End If
        End With
    End If
    
On Error GoTo errTrap
    For Each tdf In dbs.TableDefs
        If Len(tdf.Connect) > 1 Then 'Only relink linked tables
            If tdf.Connect <> ";DATABASE=" & lnkDatabase Then 'only relink tables if the are not linked right
                If Left(tdf.Connect, 4) <> "ODBC" And Left(tdf.Connect, 3) <> "WSS" Then 'Don't want to relink any ODBC tables
                    strTable = tdf.Name
                    'dbs.TableDefs(strTable).Connect = "MS Access;PWD=" & DBPassword & ";DATABASE=" & LnkDataBase
                    dbs.TableDefs(strTable).Connect = ";DATABASE=" & lnkDatabase
                    dbs.TableDefs(strTable).RefreshLink
                    log tdf.Name & " refreshed.", "Util.relinkTables"
                End If
            End If
        End If
    Next tdf
    
    Set rs = dbs.OpenRecordset("lclver")
    With rs
        .edit
        !backend = lnkDatabase
        .update
    End With
    
    DoEvents
    
    log "Table links re-synced.", "Util.relinkTables"
    
funcExit:
    relinkTables = True
    Exit Function
errTrap:
    If err = 52 Then Resume Next
    errHandler err, Error$, "Util.relinkTables"
    Resume Next
showErr:
    errHandler err, Error$, "Util.relinkTables"
    MsgBox "Invalid AeroStat Backend format.", vbCritical, "AeroStat"
    GoTo show
End Function

Public Function serialDate(d As Date) As Date
    serialDate = DateSerial(Year(d), Month(d), Day(d))
End Function

Public Function noBreaks(ByVal s As String, Optional ByVal rpl As String) As String
    s = Trim(s)
    noBreaks = Replace(s, vbCrLf, Nz(rpl, " "))
End Function

Public Function break() As String
break = vbCrLf
End Function

Public Sub errHandler(err As Integer, msg As String, Optional frm As String)
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
    setting = DLookup(v, "settings")
End Function

Public Function appendUser(ByVal usr As String, ByVal fld As String, ByVal v As Variant)
On Error GoTo errTrap
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & usr & "';")
With rs
    If .EOF Then Exit Function
    .edit
    .Fields(fld) = v
    .update
    .Close
End With
Set rs = Nothing

sexit:
    Exit Function
errTrap:
    errHandler err, Error$, "Util.appendUser"

End Function

Public Function syncTrafficLog(Optional ByVal recID As Integer, Optional ByVal tbl As String, Optional ByVal newrec As Boolean, Optional m As String)
On Error Resume Next
Dim rs As DAO.Recordset
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
                rsAlert!msg = m
            End Select
        End With
        .update
        .Bookmark = .LastModified
    End With
    
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
'    With rs
'        .edit
'        !frmtrafficlogalert = rsAlert!ID
'        .update
'    End With
End If

Set rs = CurrentDb.OpenRecordset("settings")

    rs.edit
    rs!frmTrafficLogSync = t
    rs.update
    Set rs = CurrentDb.OpenRecordset("lclver")
    rs.edit
    rs!frmTrafficLogSync = t
    rs.update
    rs.Close

Set rs = Nothing
End Function

Public Function findParentByTail(ByVal Tail As String, ByVal dir As Integer) As Integer
'dir = 1 - in, 2 - out, 3 - local
findParentByTail = 0

Dim rs As DAO.Recordset
    Select Case dir
        Case 1
            Set rs = CurrentDb.OpenRecordset("qInbound")
        Case 2
            Set rs = CurrentDb.OpenRecordset("qOutbound")
        Case 3
            Set rs = CurrentDb.OpenRecordset("qLocal")
    End Select
    
    With rs
    Do While Not .EOF
        If !Tail = Tail Then
            findParentByTail = !ID
            Exit Function
        End If
        .MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
End Function

Public Function isShiftClosed(ByVal user As String) As Integer
isShiftClosed = 0
Dim result As Boolean
result = Nz(DLookup("closed", "tblshiftmanager", "shiftID = " & Nz(DLookup("lastShift", "tbluserauth", "username = '" & user & "'"), 0)), True)
If Not result Then isShiftClosed = DLookup("lastShift", "tbluserauth", "username = '" & user & "'")
End Function

Public Function getLogName(ByVal oi As String) As String
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE opInitials = '" & oi & "';")
If rs.RecordCount = 0 Then
    getLogName = ""
    Exit Function
End If

With rs
    getLogName = !rank & " " & Left(!firstName, 1) & ". " & !lastName & "/" & !opInitials
    .Close
End With
End Function

Public Function fullName(ByVal oi As String) As String
fullName = DLookup("rank", "tblUserAuth", "opInitials = '" & oi & "'") & " " & DLookup("firstName", "tblUserAuth", "opInitials = '" & oi & "'") & " " & DLookup("lastName", "tblUserAuth", "opInitials = '" & oi & "'")
End Function

Public Function getPos(ByVal perm As Integer) As String
    Select Case perm
        Case 99
            getPos = ""
        Case 1
            getPos = "(AFM)"
        Case 2
            getPos = "(DAFM)"
        Case 3
            getPos = "(NAMO)"
        Case 4
            getPos = "(NAMT)"
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

Public Function getDuty(ByVal perm As Integer) As String
    Select Case perm
        Case 0
            getDuty = ""
        Case 1
            getDuty = "Airfield Manager"
        Case 2
            getDuty = "Deputy Airfield Manager"
        Case 3
            getDuty = "NCOIC Airfield Management Operations"
        Case 4
            getDuty = "NCOIC Airfield Management Training"
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
Dim rs As DAO.Recordset
Dim qdf As QueryDef
getAccessSP = True

    Set qdf = CurrentDb.QueryDefs("qMissionTracker")
    qdf.Parameters("fromDay") = LToZ(Date)
    Set rs = qdf.OpenRecordset()
    
    Set rs = Nothing

getAccessSp_Exit:
    Exit Function

getAccessSP_err:
    If err = 3011 Or err = 3841 Then
        MsgBox "SharePoint connection was rejected." & vbCrLf & "Changes to flight plans will not be sent to the SharePoint." & vbCrLf & vbCrLf & "Error code: " & err, vbCritical
        getAccessSP = False
        Resume Next
    Else
        MsgBox Error$
        getAccessSP = False
        Exit Function
    End If
End Function

Public Function cnlFlight(ByVal rid As Recordset) As Boolean
Dim rs As DAO.Recordset
Dim rstSP As DAO.Recordset
Set rs = rid
bClose = True
With rs

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
        .update
        syncTrafficLog !ID, "Traffic", False
    Else
        MsgBox "Flight not found", vbInformation, "Error"
    End If
End With

End Function

'Public Function checkParking(ByVal Callsign As String, ByVal Tail As String) As Integer
'Dim rs As DAO.Recordset
'Dim stn As String
'stn = DLookup("Station", "settings")
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
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("tblBaseAcft")
isT = True

If Not Left(t, 1) = "N" And Not IsNumeric(Right(t, 1)) Then t = Left(t, Len(t) - 1)

With rs
    Do While Not .EOF
        If Right(t, 4) = Right(!Tail, 4) And !acType = acType Then
            isT = False
            Exit Do
        End If
        .MoveNext
    Loop
    .Close
End With

Set rs = Nothing
        
End Function

Public Function test(str As String) As String
test = Left(str, InStr(1, str, ",") - 1)

End Function
