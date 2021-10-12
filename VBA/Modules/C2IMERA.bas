Attribute VB_Name = "C2IMERA"
Option Compare Database
Option Explicit

Public Function getCSV(ByRef PPRs As String, Optional ByVal putInClipboard As Boolean = True, Optional ByVal literalCSV As Boolean = False) As String
'Returns CSV format
'If literalCSV, returns the format as a real CSV, instead of the C2IMERA RTF-like requirement(?)
On Error GoTo errtrap

DoCmd.OpenForm "frmloading"
With Forms!frmloading
    !pBar.Visible = False
    !loadingText.Caption = "Generating data..."
End With

Dim xlApp As New Excel.Application
Dim xlBook As Excel.Workbook
    
Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qNewC2PPR")
Dim oldSQL As String: oldSQL = qdf.sql
qdf.sql = Replace(qdf.sql, "[pPPR]", PPRs)
qdf.Execute dbFailOnError
log "Getting CSV for " & qdf.RecordsAffected & " record(s)", "C2IMERA.getCSV"

Dim dataObj As New MSForms.DataObject
Dim csv() As String, header As String, idx As Integer
Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("C2IMERA_PPR")
Dim recCount As Integer

    With RS
        recCount = qdf.RecordsAffected
        ReDim Preserve csv(0 To recCount) As String
        
        Dim h: For Each h In .Fields
            If h.Name <> "ID" And h.Name <> "ATD" And h.Name <> "ATA" Then header = IIf(Nz(header) = "", "", header & "`") & h.Name
        Next h
        
        Do While Not .EOF
            Dim f: For Each f In .Fields
                'If Not Nz(f.Value) = "" Then
                
                    Select Case f.Name
                    Case "Date"
                        'csv = IIf(Nz(csv) = "", "", csv & "|") & Format(f.Value, "dd mmm yy")
                        csv(idx) = Nz(csv(idx)) & Format(f.Value, "dd mmm yy")
                        
                    Case "ETD", "ETA"
                        csv(idx) = IIf(Nz(csv(idx)) = "", "", csv(idx) & "`") & IIf(Nz(f.Value) = "", "", Format(LToZ(CDate(f.Value)), "dd mmm yy // hhnn"))
                    
                    Case "ATD", "ATA"
                        
                    Case "Aircraft Parking Location"
                    csv(idx) = IIf(Nz(csv(idx)) = "", "", csv(idx) & "`")
                    
                    Case Is <> "ID"
                        csv(idx) = IIf(Nz(csv(idx)) = "", "", csv(idx) & "`") & Replace(Nz(f.Value), vbCrLf, " | ")
                    End Select
                'End If
            Next f
            
            .delete
            If Not .EOF Then
                .MoveNext
                If Not .EOF Then idx = idx + 1
            End If
        Loop
        
        .Close
        Set RS = Nothing
    End With

    Dim result
    Select Case literalCSV
    Case True
        result = Replace("""" & header & """" & vbCrLf & """" & UCase(csv) & """", "|", """,""")
    Case False
        result = Replace(header & vbCrLf & UCase(join(csv, vbCrLf)), "`", vbTab)
    End Select
    
    getCSV = result
    dataObj.SetText result
    dataObj.putInClipboard
    
    On Error GoTo xlErr
    
    Set xlBook = xlApp.Workbooks.add
    xlBook.Activate
    'xlApp.Visible = True
    With xlBook
        Dim sheet As Excel.Worksheet: Set sheet = .ActiveSheet
        With sheet.range("A1")
            .ColumnWidth = 50
            .PasteSpecial
        End With

        sheet.range("A1:AE" & 1 + recCount).copy
        Set sheet = Nothing
    End With
    
    On Error Resume Next: DoCmd.Close acForm, "frmloading", acSaveNo
    
    MsgBox "C2IMERA PPR format generated successfully!" & vbCrLf & vbCrLf & _
            "Paste the data to the C2IMERA window first, then press ""OK""", vbInformation, "C2IMERA"
xlExit:
On Error GoTo errtrap
    xlBook.Close False
    xlApp.Quit
fexit:
    qdf.sql = oldSQL
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set qdf = Nothing
    Set dataObj = Nothing
    Util.trunc "C2IMERA_PPR", True
    Exit Function
    
xlErr:
    ErrHandler err, "Excel ran into an error: " & Error$, "C2IMERA.getCSV"
    GoTo xlExit
    
errtrap:
    ErrHandler err, Error$, "C2IMERA.getCSV"
    GoTo fexit
    Resume Next

End Function

Private Function readClipboard() As Variant
Dim dataObj As New MSForms.DataObject
dataObj.GetFromClipboard

readClipboard = dataObj.GetText

dataObj.clear
Set dataObj = Nothing
End Function

Public Function getFuelInPounds(ByVal amount As Variant, Optional ByVal fuelType As String = "JP8/Jet A") As Double
Dim nbr As Double, FACTOR As Double
amount = Trim(Replace(amount, ",", ""))

'Isolate the number from the variable
Dim i: For i = 1 To Len(amount)
    If IsNumeric(Mid(amount, i, 1)) Then
        nbr = nbr & Mid(amount, i, 1)
    Else
        Exit For
    End If
Next i

'If you wanna add support to convert different fuel types, define them here
Select Case fuelType
    Case "JP8/Jet A"
        FACTOR = 6.8 'The weight (in pounds) of 1 US gallon of JP8 at STP
        
    'Case "Some other fuel type"
    
End Select

'Do the math
If amount Like "*GAL*" Or amount Like "*GALS*" Or amount Like "*G" Then
    nbr = nbr * FACTOR
ElseIf amount Like "*K" Then
    nbr = nbr * 1000
End If

'Result
getFuelInPounds = nbr

End Function

Public Function getC2FormatX(ByVal pprID As Integer) As String
'Returns CSV Format record.
Dim rs1 As DAO.Recordset: Set rs1 = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE ID = " & pprID)
Dim rs2 As DAO.Recordset: Set rs2 = CurrentDb.OpenRecordset("C2IMERA_PPR")
Dim c2Headers As Variant: c2Headers = Array("ID,Date,Callsign,Aircraft MDS,tail,depPoint,ATA,ETA,spot,details,ETD,ATD,Destination,dvCode,POC,fuel,pax,hazCargo,customs,fleetSvc,lodgingInfo,crewTrans,deployment,PPR,portToPort,persco,SAAM,msn")

    rs2.AddNew
    Dim fld: For Each fld In rs1.Fields
        rs2.Fields(fld.Name) = fld.Value
    Next fld
    rs2.update
    rs1.MoveNext
    
    'Open DLookup("dbroot", "tblsettings") & "C2.csv" For Output As #1
        
        'Get column heads for CSV
        Dim result() As String
        Dim idx As Integer
        Dim f: For Each f In rs2.Fields
            ReDim Preserve result(0 To idx) As String
            result(idx) = f.Name
            idx = idx + 1
        Next f
        
        'Print column heads as first line
        'Print #1, join(result, ",")
        getC2FormatX = join(result, ",")
        
        'Print records to be copied
        
        Dim qdf: For Each qdf In CurrentDb.QueryDefs
            If Left(qdf.Name, 1) = "q" Then
                log "Exporting query: " & qdf.Name, "Util.exportSchema"
                'Application.ExportXML acExportQuery, qdf.Name, , Path & qdf.Name & ".xsd"
                'CurrentDb.Execute "INSERT INTO [@SQL] (qName,[SQL]) SELECT '" & qdf.Name & "', '" & qdf.sql & "'", dbFailOnError
                'EXPORT METHOD
                'Print #1, """" & qdf.Name & """,""" & Replace(qdf.sql, """", "'") & """"
                getC2FormatX = getC2FormatX & vbCrLf & """" & qdf.Name & """,""" & Replace(qdf.sql, """", "'") & """"
                DoEvents
            End If
        Next
    'Close #1

End Function

Public Function importPPR() As Boolean
Dim rsFrom As DAO.Recordset: Set rsFrom = CurrentDb.OpenRecordset("tblPPR")
Dim rsTo As DAO.Recordset

End Function
