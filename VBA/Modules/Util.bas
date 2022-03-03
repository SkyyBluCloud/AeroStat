Attribute VB_Name = "Util"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_SYNC = &H0
Public Const SND_LOOP = &H8

Option Compare Database

Public Function createBackend() As Variant
On err GoTo errtrap
Dim fd As Office.FileDialog: Set fd = Access.FileDialog(msoFileDialogSaveAs)
Dim fso As New FileSystemObject
Dim f, fr, fc
Dim acc As New Access.Application
Dim saveLocation As String

    With fd 'File dialog for save location
        .title = "Save New Backend"
        .InitialFileName = "ICAO DATA.accdb"
        '.Filters.add "Access Database", "*.accdb"
        If .show Then
            Dim s: For Each s In .SelectedItems
                saveLocation = s
            Next
            'fso.DeleteFile saveLocation
        Else
            'Cancelled by user
            GoTo sexit
        End If
    End With
    
    With acc 'Create the new database file
        Dim schema As String
        
        If fso.FileExists(saveLocation) Then fso.DeleteFile (saveLocation)
        .DBEngine.CreateDatabase saveLocation, dbLangGeneral
        
        Dim subDB As DAO.Database
        .OpenCurrentDatabase (saveLocation)
        Set subDB = .CurrentDb
        
        schema = .CurrentProject.Path & "\Schema\"
        If Not fso.FolderExists(schema) Then
            Set fd = Access.FileDialog(msoFileDialogFolderPicker)
            With fd
                .title = "Select schema folder"
                If .show Then
                    Dim S1: For Each S1 In .SelectedItems
                        schema = S1
                    Next
                Else
                    'Cancelled by user
                    GoTo sexit
                End If
            End With
                
        Else
            
        End If
        
        Set fr = fso.GetFolder(schema)
        Set fc = fr.Files
        
        On Error GoTo runtimeErr
        For Each f In fc 'Lookup each file in schema folder, then import
            Select Case Right(f.Name, 4)
            Case ".xsd"
                log "Creating schema from XML: " & f.Name, "frmSetup.btnCreateBackend"
                .ImportXML f, acStructureOnly
            Case ".xml"
                If Not f.Name = "@DEBUG.xml" Then
                    log "Importing table from XML: " & f.Name, "frmSetup.btnCreateBackend"
                    .ImportXML f, acStructureAndData
                End If
            Case ".csv"
                log "Importing CSV.........", f.Name & ".createBackend"
                Dim importTable As TableDef: Set importTable = subDB.CreateTableDef("qryImport")
                With importTable
                    .Fields.Append .CreateField("name", dbText, 255)
                    .Fields.Append .CreateField("sql", dbMemo, 65535)
                    
                End With
                subDB.TableDefs.refresh
                
                .DoCmd.TransferText , , "qryImport", f, True
                
                Dim RS As DAO.Recordset: Set RS = subDB.OpenRecordset("qryImport")
                With RS: Do While Not .EOF
                    subDB.CreateQueryDef !Name, !sql
                    log !Name, f.Name & ".createBackend"
                    .MoveNext
                Loop: End With
                subDB.QueryDefs.refresh
                
                subDB.TableDefs.delete ("qryImport")
            End Select
        Next
        On Error GoTo errtrap
        
        .CloseCurrentDatabase
        .Quit acQuitSaveAll
    End With
    
    createBackend = saveLocation
    
'    If Util.relinkTables(saveLocation) Then
'        log "Done!", "frmSetup.btnCreateBackend"
'        tabCtl = tabCtl + 1
'    Else
'        log "Something went wrong.", "frmSetup.btnCreateBackend_Click", "ERR"
'        GoTo sexit
'    End If

sexit:
    'Cleanup
    Set acc = Nothing
    Set fd = Nothing
    Set fso = Nothing
    Set fr = Nothing
    Set fc = Nothing
    Exit Function
errtrap:
    Select Case err
    Case 31550
        MsgBox Error$ & " (" & err & ")", vbCritical, "Error"
        Resume Next
    Case 76
        MsgBox Error$ & " (" & err & ")", vbCritical, "Error"
    End Select
    Resume sexit
runtimeErr:
    ErrHandler err, Error$, "Util.createBackend"
    Resume Next
End Function

Public Sub printReport(ByVal rName As String, Optional ByVal openargs As Variant)
On Error GoTo errtrap
If Not CurrentProject.AllReports(rName).IsLoaded Then
    DoCmd.OpenReport rName, acViewPreview, , , acHidden, openargs
End If
    DoCmd.SelectObject acReport, rName
    DoEvents
    
    DoCmd.RunCommand acCmdPrint
    
sexit:
DoCmd.Close acReport, rName
Exit Sub
errtrap:
Select Case err
    Case Is <> 2501
        ErrHandler err, Error$, "Util" & ".printReport"
End Select
Resume sexit
End Sub

Sub qDefs()
On Error GoTo errtrap
Dim key As String: key = "getUSN"
Dim qdf: For Each qdf In CurrentDb.QueryDefs
    If Left(qdf.Name, 1) = "q" And InStr(1, qdf.sql, key) <> 0 Then
        log qdf.Name, "qDefs"
    End If
Next
sexit:
Exit Sub
errtrap:
ErrHandler err, Error$, "qDefs"
End Sub

Public Function getUser() As Variant
' This procedure uses the Win32API function util.getuserName
' to return the name of the user currently logged on to
' this machine. The Declare statement for the API function
' is located in the Declarations section of this module.
   
    Dim strBuffer As String
    Dim lngSize As Long
        
    strBuffer = String(100, " ")
    lngSize = Len(strBuffer)
    
    If GetUserName(strBuffer, lngSize) = 1 Then
        getUser = Left(strBuffer, lngSize - 1)
    End If
    
End Function

Public Function getWorkstation() As Variant
    Dim strBuffer As String
    Dim lngSize As Long
        
    strBuffer = String(100, " ")
    lngSize = Len(strBuffer)

    If GetComputerName(strBuffer, lngSize) = 1 Then
        getWorkstation = Left(strBuffer, lngSize)
    End If

End Function


Public Function getSettings(key As String) As Variant

getSettings = DLookup("data", "tblSettings", "key = '" & key & "'")
End Function

Public Sub clearConnections()
On Error GoTo errtrap
    For Each t In CurrentDb.TableDefs
        t.Connect = ";"
    Next t
    log "Done!", "Util.saveConnections"
    
fexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, "Util.clearConnections"
    Resume Next
End Sub

Public Sub saveConnections()
On Error GoTo errtrap
    trunc "tblConnectionStrings", True
    For Each t In CurrentDb.TableDefs
        If Len(t.Connect) > 1 Then 'Only relink linked tables
            If Left(t.Connect, 4) <> "ODBC" And Left(t.Connect, 3) <> "WSS" Then
'                CurrentDb.Execute "INSERT INTO tblConnectionStrings (table, connect) VALUES (""" & t.Name & """, """ & t.Connect & """)"
                Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("tblConnectionStrings")
                If t.Name Like "tbl*" Then
                    With RS
                        .AddNew
                        !Table = t.Name
                        !Connect = t.Connect
                        .update
                    End With
                End If
            End If
        End If
    Next t
    log "Done!", "Util.saveConnections"
    
fexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, "Util.saveConnections"
End Sub

Public Sub closeAllForms()
On Error Resume Next
    For Each f In CurrentProject.AllForms
        If f.IsLoaded Then
            DoCmd.Close acForm, f.Name
        End If
    Next
End Sub

Public Function GetHTTPResponse(URL As String) As String
Dim msXML As New MSXML2.XMLHTTP60
With msXML
  .Open "Get", URL, False
  .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
  .Send
  GetHTTPResponse = .responseText
End With
Set msXML = Nothing
End Function


Public Function doCAPS(frm As Form) As String
On Error Resume Next
'Triggers on TextBox Exit
For Each ctl In frm.Controls
    If TypeOf ctl Is TextBox Then
        'If ctl.Value <> Replace(Nz(ctl.DefaultValue), """", "") Then ctl.Value = UCase(ctl.Value)
        If ctl.Value <> Replace(Nz(ctl.DefaultValue), """", "") Then ctl.Value = UCase(ctl.Value)
    End If
Next

End Function

Public Function getTime4Char(ByVal char4 As String) As Variant
'Converts a 4-digit string, representing a time, into a time format Access can understand.
' t <string>: 4 digit time
On Error GoTo errtrap

    If Not IsNumeric(char4) Or _
            Len(char4) > 4 Or _
            Right(char4, 2) > 59 Or _
            (Len(char4) = 4 And Left(char4, 2) > 23) Then
            
        getTime4Char = Null
        Exit Function
    End If
    
    char4 = Format(char4, "0000")
    getTime4Char = TimeSerial(Left(char4, 2), Right(char4, 2), 0)
    
fexit:
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.getTime4Char"
End Function

Public Sub lclver()
'Lazy way to open the lclver table.
    DoCmd.OpenTable "lclver"
End Sub

'Saves table relationships and saves them to a table to be exported later.
'dbPath <string> = The path of the database to run the function on. Default: DLookup("backend", "tblsettings")
Public Function saveRelations(Optional ByVal dbPath As String)
If Nz(dbPath) = "" Then dbPath = DLookup("data", "tblSettings", "key = ""backend""")
Dim acc As New Access.Application
Dim db As DAO.Database: Set db = CurrentDb
Dim RS As DAO.Recordset: Set RS = db.OpenRecordset("SELECT * FROM [@RELATIONS] IN '" & dbPath & "'", , dbFailOnError)
    
    log "Clearing old relations...", "Util.saveRelations"
    db.Execute "DELETE * FROM [@RELATIONS] IN '" & dbPath & "'", dbFailOnError
    
    For Each r In db.Relations
        With r
            sName = Mid(.Name, InStr(1, .Name, "].") + 2)
            log "Saving relationship: " & sName, "Util.saveRelations"
            
            sfields = ""
            Dim f: For Each f In .Fields
                If Nz(sfields) = "" Then
                    sfields = f.Name
                Else
                    sfields = sfields & ";" & f.Name
                End If
                
            Next: With RS
                .AddNew
                !rName = sName
                !Attributes = r.Attributes
                !ptable = r.Table
                !ftable = r.ForeignTable
                !Fields = sfields
                .update
            End With
            
            'CurrentDb.Execute "INSERT INTO [" & DLookup("backend", "tblsettings") & "].[@RELATIONS] (name, attributes, pTable, fTable, fields) SELECT '" & _
                                sName & "' AS newName, '" & .Attributes & "' AS newAttributes, '" & .Table & "' AS newTable, '" & _
                                .ForeignTable & "' AS newForeignTable, '" & sfields & "' AS newFieldString IN;", dbFailOnError
            
        End With
    Next
    
db.Close
Set db = Nothing
End Function

'Don't do it.
Public Sub truncAll(ByVal dbPath As Variant)
On Error Resume Next
    If MsgBox("THIS IS DANGEROUS! TURN BACK NOW!", vbCritical + vbYesNo, "Truncate (Thats fancy for DELETE EVERYTHING)") = vbNo Then
    If MsgBox("THIS CANNOT BE UNDONE! CANCEL THIS IMMEDIATELY!", vbCritical + vbYesNo, "Truncate (Thats fancy for DELETE EVERYTHING)") = vbNo Then
    If MsgBox("SERIOUSLY? CANCEL THIS PROCESS? PLEASE?", vbCritical + vbYesNo, "Truncate (Thats fancy for DELETE EVERYTHING)") = vbNo Then
        Dim db As DAO.Database
        Dim app As New Access.Application: app.OpenCurrentDatabase dbPath
        Set db = app.CurrentDb
    
        Dim tdf: For Each tdf In db.TableDefs
            If Left(tdf.Name, 3) = "tbl" Then
                db.Execute "DELETE * FROM " & tdf.Name, dbFailOnError
            End If
        Next
    End If
    End If
    End If

    Set db = Nothing
    app.CloseCurrentDatabase
    app.Quit
    DoEvents
    Set app = Nothing
    DoEvents
    MsgBox "Done. It's like it never happened...", vbInformation, "Trunc"
End Sub

'Don't do this either.
Public Sub trunc(ByVal tbl As String, Optional ByVal surpressWarning As Boolean)
Dim db As DAO.Database: Set db = CurrentDb
    If Not surpressWarning Then
        If MsgBox("THIS IS DANGEROUS! TURN BACK NOW!", vbCritical + vbYesNo, "Truncate (Thats fancy for DELETE EVERYTHING)") = vbYes Then
            Exit Sub
        End If
    End If
    Dim timeStart As Date: timeStart = Now
    
    db.Execute "DELETE * FROM " & tbl, dbFailOnError
    DoEvents
    
    log "Truncated " & tbl & " in " & DateDiff("s", timeStart, Now) & " seconds. (" & db.RecordsAffected & " record(s) )", "Util.trunc"
    Set db = Nothing
End Sub

Public Sub exportSchema(Optional Path As Variant = Null)
On Error GoTo errtrap
Dim tdf As DAO.TableDef
If IsNull(Path) Then Path = CurrentProject.Path & "\DB EXPORT\Schema\"
'Path = Replace(Path, "%USERPROFILE%", Replace(Environ$("userprofile"), "C:\", "D:\"))

    log "Creating path...", "Util.exportSchema"
    log IIf(Util.createPath(Path), "[" & Path & "] Success!", "Failed to create path."), "Util.exportSchema"

    saveRelations
    
    Open Path & "SQL.csv" For Output As #1
    
        Print #1, """Name"",""SQL"""
        
        Dim qdf: For Each qdf In CurrentDb.QueryDefs
            If Left(qdf.Name, 1) = "q" Then
                log "Exporting query: " & qdf.Name, "Util.exportSchema"
                'Application.ExportXML acExportQuery, qdf.Name, , Path & qdf.Name & ".xsd"
                'CurrentDb.Execute "INSERT INTO [@SQL] (qName,[SQL]) SELECT '" & qdf.Name & "', '" & qdf.sql & "'", dbFailOnError
                'EXPORT METHOD
                Print #1, """" & qdf.Name & """,""" & Replace(qdf.sql, """", "'") & """"
                DoEvents
            End If
        Next
    Close #1
    
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.Name, 3) = "tbl" Then
            log "Exporting schema: " & tdf.Name, "Util.exportSchema"
            'Application.ExportXML acExportTable, tdf.Name, , Path & tdf.Name & ".xsd"
            Application.ExportXML acExportTable, tdf.Name, , Path & tdf.Name & ".xsd"
            DoEvents
        ElseIf Left(tdf.Name, 1) = "@" Then
            log "Exporting table: " & tdf.Name, "Util.exportSchema"
            Application.ExportXML acExportTable, tdf.Name, Path & tdf.Name & ".xml"
            DoEvents
        End If
    Next
    
sexit:
    log "Done! Schema located at " & Path, "Util.exportSchema"
    Exit Sub
errtrap:
    ErrHandler err, Error$, "Util.exportSchema"
End Sub

Public Function fixCase(ByRef s As Control)
On Error Resume Next
    s.Value = UCase(Left(s.Value, 1)) & Right(LCase(s.Value), Len(s.Value) - 1)
End Function

Public Function cETA(ByVal DOF As Variant, ByVal ETD As Variant, ByVal ETE As Variant, _
                    Optional ByVal ETA As Variant = Null, Optional ByVal ATD As Variant = Null, Optional ByVal ATA As Variant = Null) As Variant
On Error GoTo errtrap
If IsNull(DOF) Then Exit Function
'DOF = CDate(DOF): ETD = CDate(Nz(ETD, 0)): ETE = CDate(Nz(ETE, 0))
'ETA = CDate(Nz(ETA, 0))
'ATD = CDate(Nz(ATD, 0)): ATA = CDate(Nz(ATA, 0))
'[DOF]+IIf([ETA] Is Null,IIf([ATD] Is Null,[ETD],[ATD])+[ETE],[ETA])
    Select Case False
        Case IsDate(DOF), IsDate(ETD), IsDate(ETE)
            Exit Function
            
    End Select
    
    If ATA = "" Then ATA = Null
    
    Dim CTD: CTD = DateValue(DOF) + Nz(ATD, ETD)
    Dim CTA: CTA = DateValue(DOF) + Nz(ATA, Nz(ETA, Nz(ATD, ETD) + ETE))
    
    If CTA < CTD Then CTA = DateAdd("d", 1, CTA)
    
    cETA = CTA
    
'    If Not IsNull(ETA) Then
'        If DateValue(DOF + ETA) <> DateValue(DOF + Nz(ATD, ETD) + ETE) Then
'            cETA = DateValue(DOF + Nz(ATD, ETD) + ETE) + ETA
'        End If
'
'    Else
'        cETA = DateValue(DateValue(DOF) + Nz(ATD, ETD) + ETE) + TimeValue(Nz(ATA, Nz(ETA, Nz(ATD, ETD) + ETE)))
'    End If
') + TimeValue(Nz(ATA, Nz(ETA)))
fexit:
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.cETA"
    Resume Next
End Function

Function LToZ(ByVal lcl As Date) As Date
    Dim Timezone As Integer: Timezone = DLookup("data", "tblSettings", "key = ""timezone""")
    If DLookup("data", "tblSettings", "key = ""dst""") And isDST(DateValue(lcl)) Then Timezone = Timezone + 1
    'If lcl = "" Then Exit Function
    
    LToZ = DateAdd("h", -Timezone, lcl)
    'LToZ = TimeSerial(Hour(LToZ), Minute(LToZ), 0)
    
End Function

Function ZToL(ByVal zulu As Date) As Date
    Dim Timezone As Integer: Timezone = DLookup("data", "tblSettings", "key = ""timezone""")
    If DLookup("data", "tblSettings", "key = ""dst""") And isDST(DateValue(zulu)) Then Timezone = Timezone + 1
    
    'If zulu = "" Then Exit Function
    
    ZToL = DateAdd("h", Timezone, zulu)
    'ZToL = TimeSerial(Hour(ZToL), Minute(ZToL), 0)
End Function

Public Function isDST(Optional ByVal d0 As Variant = Null, Optional locale As String = "US") As Boolean
Dim dstOn As Date, dstOff As Date
d0 = Nz(d0, Date)

    Select Case locale
    Case "US"
        dstOn = DateSerial(Year(d0), 3, 8) '"#MAR 8 " & Year(d0) & "#"
        dstOff = DateSerial(Year(d0), 11, 8) '"#NOV 8 " & Year(d0) & "#"
        
    Case "EU"
        dstOn = DateSerial(Year(d0), 3, 31) '"#MAR 31 " & Year(d0) & "#"
        dstOff = DateSerial(Year(d0), 11, 1) '"#NOV 1 " & Year(d0) & "#"
    End Select
    
    isDST = d0 >= NextSun(dstOn) And d0 < LastSun(dstOff)
End Function

Private Function NextSun(D1 As Date) As Date
   NextSun = IIf(Weekday(D1) = 1, D1, (D1 + 7) - (Weekday(D1) - 1))
End Function

Private Function LastSun(D1 As Date) As Date
    LastSun = D1 - IIf(Weekday(D1) = 1, 7, Weekday(D1) - 1)
End Function

'Public Sub exportSchema(Optional Path As String = "%USERPROFILE%\Documents\AeroStat\Schema\")
'Dim f As String
'Dim tdf As DAO.TableDef
'Path = Replace(Path, "%USERPROFILE%", Environ$("userprofile"))
'
'    log "Creating path...", "Util.exportSchema"
'    log IIf(Util.createPath(Path), "Success!", "Failed to create path."), "Util.exportSchema"
'
'    For Each tdf In CurrentDb.TableDefs
'        If Left(tdf.Name, 3) = "tbl" Then
'            log "Exporting schema: " & tdf.Name, "Util.exportSchema"
'            Application.ExportXML acExportTable, tdf.Name, , Path & tdf.Name & ".xsd"
'            DoEvents
'        End If
'    Next
'    log "Done!", "Util.exportSchema"
'End Sub

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

sexit:
    Exit Sub
errtrap:
    Util.ErrHandler err, Error$, "Util.EXPORT"
    Resume sexit
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

Public Function getOpInitials(ByVal username As String) As Variant
    'getOpInitials = DLookup("opInitials", "tbluserauth", "username = '" & IIf(username <> "", username, Util.getUser) & "'")
    getOpInitials = DLookup("opInitials", "tbluserauth", "username = '" & username & "'")
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
    
fexit:
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
fexit:
    Exit Function
errtrap:
    ErrHandler err, Error$, "Util.testBackend"
End Function

Public Function relinkTables(Optional ByVal backend As Variant = Null, Optional ByRef loading As Variant = Null) As Boolean
On Error GoTo errtrap
Const DO_LOG As Boolean = False
Dim db As DAO.Database
Dim RS As DAO.Recordset
Dim tdf As DAO.TableDef
Dim ld As Boolean
Dim strTable, lnkDatabase, lnkAtlas, newFile, key As String
Set db = CurrentDb

If IsNull(loading) Then
    DoCmd.OpenForm "frmLoading"
    Set loading = Forms!frmloading
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
    lnkDatabase = Nz(DLookup("data", "tblSettings", "key = ""backend"""), _
                        Nz(DLookup("data", "tblSettings", "key = ""backend""")) _
                    )
Else
    lnkDatabase = backend
End If
lnkAtlas = Nz(DLookup("data", "tblSettings", "key = 'atlas'"))
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
                If DO_LOG Then log "Cancelled by user.", "Util.relinkTables"
                GoTo fexit
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
                    If Left(tdf.Connect, 4) <> "ODBC" And Left(tdf.Connect, 3) <> "WSS" Then 'Don't want to relink any ODBC or SP tables
                        strTable = tdf.Name
                        'db.TableDefs(strTable).Connect = "MS Access;PWD=" & DBPassword & ";DATABASE=" & LnkDataBase
                        If strTable Like "tbl*" Then
                            db.TableDefs(strTable).Connect = ";DATABASE=" & lnkDatabase
                            db.TableDefs(strTable).RefreshLink
                        End If
                        
                        If DO_LOG Then log tdf.Name & " refreshed.", "Util.relinkTables"
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
    If DO_LOG Then log "Table links re-synced.", "Util.relinkTables"
    loading!loadingText.Caption = "Tables Updated."
    DoEvents
    db.Execute "UPDATE tblSettings SET data = '" & lnkDatabase & "' WHERE key = 'backend'"
    db.Execute "UPDATE lclver SET backend = '" & lnkDatabase & "'"
    relinkTables = True
    
fexit:
    If ld Then DoCmd.Close acForm, "frmLoading"
    Call saveConnections
    Exit Function
errtrap:
    If err = 52 Then Resume Next
    If DO_LOG Then ErrHandler err, Error$, "Util.relinkTables"
    Resume Next
showErr:
    If DO_LOG Then ErrHandler err, Error$, "Util.relinkTables"
    MsgBox "Invalid AeroStat Backend format.", vbCritical, "AeroStat"
    GoTo show
End Function

Public Function serialDate(d As Date) As Date
    serialDate = DateSerial(Year(d), Month(d), Day(d))
End Function

Public Function noBreaks(ByVal s As String, Optional ByVal separator As String) As String
    s = Trim(s)
    noBreaks = Replace(s, vbCrLf, Nz(separator, ""))
End Function

Public Function strBreak() As String
break = vbCrLf
End Function

Public Sub ErrHandler(err As Integer, msg As String, Optional frm As String)
'    Debug.Print Format(Now, "dd-mmm-yy hh:nnL ") & "(" & err & ") " & IIf(IsNull(frm), "", "[" & frm & "] ") & msg
    log "(" & err & ") " & msg, IIf(IsNull(frm), "", frm & " "), "ERR"
    'Beep
End Sub

Public Sub log(msg As String, module As String, Optional priority As String = "INFO")
On Error GoTo errtrap
Dim db As DAO.Database: Set db = CurrentDb
Dim sql As String:  sql = "INSERT INTO [@DEBUG] (username,initials,computername,priority,module,details) " & _
                            "SELECT tblUserAuth.username, tblUserAuth.opInitials, tblUserAuth.lastsystem, '" & priority & "', '" & module & "', " & """" & msg & """" & _
                            " FROM tblUserAuth WHERE tblUserAuth.username = '" & Util.getUser & "';"

    db.Execute sql, dbFailOnError
    Debug.Print Format(Now, "dd-mmm-yy hh:nn:ssL") & "[" & priority & "] " & module & ": " & msg
    
    If CurrentProject.AllForms("CONSOLE").IsLoaded Then Forms!CONSOLE.update
sexit:
Exit Sub
errtrap:

Resume Next
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
    setting = DLookup("data", "tblSettings", "key = """ & v & """")
End Function

Public Function appendUser(ByVal usr As String, ByVal fld As String, ByVal v As Variant)
On Error GoTo errtrap
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & usr & "';")
With RS
    If .EOF Then Exit Function
    .edit
    .Fields(fld) = v
    .update
    .Close
End With
Set RS = Nothing

sexit:
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
        !opInitials = DLookup("opinitials", "tbluserauth", "username = '" & Util.getUser & "'")
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
        .update
        .Bookmark = .LastModified
    End With
    
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Util.getUser & "'")
'    With rs
'        .edit
'        !frmtrafficlogalert = rsAlert!ID
'        .update
'    End With
End If

Set RS = CurrentDb.OpenRecordset("tblSettings")

    RS.edit
    RS!frmTrafficLogSync = t
    RS.update
    Set RS = CurrentDb.OpenRecordset("lclver")
    RS.edit
    RS!frmTrafficLogSync = t
    RS.update
    RS.Close

Set RS = Nothing
End Function

Public Function findParentByTail(ByVal Tail As String, ByVal dir As Integer) As Integer
'dir = 1 - in, 2 - out, 3 - local
findParentByTail = 0
If Tail = "" Then Exit Function
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
            getDuty = "Deputy/Assistant Airfield Manager"
            If alt Then getDuty = "Assistant Airfield Manager"
        Case 3
            getDuty = "NCOIC/Manager, Airfield Management Operations"
            If alt Then getDuty = "Airfield Management Operations Manager"
        Case 4
            getDuty = "NCOIC/Manager Airfield Management Training"
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
'stn = DLookup("data","tblSettings","key = 'station'")
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
