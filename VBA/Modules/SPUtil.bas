Attribute VB_Name = "SPUtil"
Option Compare Database
Option Explicit
Const spTable As String = "spPPRLog"

Public Function getSPID(Optional ByVal username As Variant = Null) As Variant
username = Nz(username, Util.getUser)
Dim fName As String: fName = DLookup("firstName", "tblUserAuth", "username = '" & username & "'")
Dim lName As String: lName = DLookup("lastName", "tblUserAuth", "username = '" & username & "'")

    getSPID = DLookup("ID", "spUserInfo", "[First name] = """ & UCase(fName) & """ AND [Last name] = """ & UCase(lName) & """")
    
End Function

Public Function getSPField(ByVal solution As String) As Variant

    getSPField = DLookup("SPField", "tblSPConversion", "solution = '" & Eval(solution) & "'")

'    Select Case lclFld
'        Case "PPR": getSPField = "[PPR #]"
'        Case "arrDate": getSPField = "[Date] + [ETA (L)]"
'        Case "Callsign": getSPField = "[C/S]"
'        Case "Type": getSPField = "Acft Type"
'        Case "depDate": getSPField = ""
'        Case "ETD (L)": getSPField = "format(datevalue(arrDate),""dd"") & ""/"" & TimeValue(depDate)"
'        Case "depPoint": getSPField = "From"
'        Case "Destination": getSPField = "To"
'        Case "Remarks": getSPField = "Purpose + Remarks"
'        Case Else: getSPField = Null
'    End Select
    
End Function

Public Function getLCLField(ByVal spFld As String) As Variant

    

    Select Case spFld
        Case "PPR #": getLCLField = "PPR"
        Case "Date": getLCLField = "DateValue(arrDate)"
        Case "C/S": getLCLField = "Callsign"
        Case "Acft Type": getLCLField = "Type"
        Case "ETA (L)": getLCLField = "TimeValue(arrDate)"
        Case "ETD (L)": getLCLField = "format(datevalue(arrDate),""dd"") & ""/"" & TimeValue(depDate)"
        Case Else: getLCLField = Null
    End Select
    
End Function

Public Function updateSP(ByVal PPR As String, Optional ByVal newrec As Variant = Null) As Variant
Dim db As DAO.Database: Set db = CurrentDb


End Function
