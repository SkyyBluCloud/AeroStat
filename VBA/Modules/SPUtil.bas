Attribute VB_Name = "SPUtil"
Option Compare Database

Public Function getSPName(ByVal spID As Integer) As String
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE spID = " & spID & ";")
If RS.RecordCount = 0 Then
    getSPName = ""
    Exit Function
End If

With RS
    getSPName = !rankID & " " & Left(!firstName, 1) & ". " & !lastName & "/" & !opInitials
    .Close
End With
End Function

Public Function syncPPRs(ByRef rsPPR As DAO.Recordset) As Boolean
Dim rsSP As DAO.Recordset
Dim qdf As DAO.QueryDef

    With rsPPR: Do While Not .EOF
        qdf.Parameters("mtbid") = !spID
        Set rsSP = qdf.OpenRecordset
        With rsSP
        
            ![Start Time] = Nz(![Start Time])
            ![End Time] = Nz(![End Time])
            ![PPR #] = formPPR
            ![Call Sign] = Callsign
            ![Aircraft Type] = rsPPR!Type
            
            If Not ![Tail Number] Like rsPPR!Tail Then
                If Nz(![Tail Number]) = "" Then
                    ![Tail Number] = rsPPR!Tail
                ElseIf MsgBox("Tail number does not match SharePoint:" & vbCrLf & "SharePoint Tail: " & Nz(![Tail Number], "None") & vbCrLf & "Your Tail: " & Tail & _
                        vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
                   ![Tail Number] = rsPPR!Tail
                Else
                    rsPPR!Tail = ![Tail Number]
                End If
            End If
            
            ![Current ICAO] = depPoint
            If Not rsFP.EOF Then
                ![ETA (Z)] = LToZ(rsFP!arrDate)
            Else
                ![ETA (Z)] = LToZ(arrDate)
            End If
            ![Next ICAO] = pprDestination
            ![ETD (Z)] = LToZ(Nz(depDate))
            ![DV Code] = dvCode
            If Status = "Pending" And Not Remarks Like "*PENDING APPROVAL*" Then
                Remarks = "*PENDING APPROVAL*" & vbCrLf & Remarks
            End If
            .Fields("Details") = Remarks
            !Fuel = Fuel
            '![Parking Spot/Location] = Spot
            
            '''''Parking''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Select Case Nz(Spot)
                Case "", "AMC", "TBD"  'We do not have assignment
                    Select Case Nz(![Parking Spot/Location])
                    Case "", "AMC", "TBD"
                        If Spot <> Nz(![Parking Spot/Location]) Then ![Parking Spot/Location] = Spot
                    Case Else
                        GoTo els
                    End Select
                    
                    Spot = IIf(Left(Nz(![Parking Spot/Location]), 3) = "HOT", "HC" & Right(Nz(![Parking Spot/Location]), 1), Nz(![Parking Spot/Location]))
                    '![Parking Spot/Location] = Spot
                        
                Case Is <> ![Parking Spot/Location] 'We have assignment, but it doesnt match the SharePoint
els:
                    If Nz(![Parking Spot/Location]) = "" Then
                    
                        ![Parking Spot/Location] = Spot
                        
                    ElseIf MsgBox("Parking assignment does not match SharePoint:" & vbCrLf & "SharePoint Spot: " & ![Parking Spot/Location] & vbCrLf & "Your Spot: " & Spot & _
                    vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
                        ![Parking Spot/Location] = Spot
                    Else
                        Spot = ![Parking Spot/Location]
                    End If
            End Select
        End With
    Loop: End With
    
        
End Function

Public Function syncUserID()
Dim rsSP As DAO.Recordset
Dim RS As DAO.Recordset
Dim N As Integer
Set RS = CurrentDb.OpenRecordset("tblUserAuth")
ErrHandler 0, "Started.", "SPUtil.syncUserID"
With RS: Do While Not .EOF
    N = N + 1
    ErrHandler 0, N & "/" & .RecordCount & " records complete...", "SPUtil.syncUserID"
    DoEvents
    Set rsSP = CurrentDb.OpenRecordset("SELECT ID FROM ADPMUsers WHERE [User name] = '" & !username & "'")
    If Not rsSP.EOF Then
        .edit
        !spID = rsSP!ID
        .Update
    End If
    rsSP.Close
    Set rsSP = Nothing
    .MoveNext
Loop: .Close: End With
Set RS = Nothing

ErrHandler 0, "DONE.", "SPUtil.syncUserID"
End Function
