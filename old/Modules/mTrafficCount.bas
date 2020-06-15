Attribute VB_Name = "mTrafficCount"
Option Compare Database

Function alpha(ByVal s As String) As String
    For i = 1 To Len(s)
        b = Mid(s, i, 1)
        Select Case b
            Case "a" To "z", "A" To "Z", " "
                alpha = alpha & b
        End Select
    Next
End Function

Function ICAOToCount(ByVal icao As String, ByVal sts As Integer) As String
'For RJTY, isOther = JSDF
Dim isComm As Boolean
Dim isMil As Boolean
Dim isother As Boolean
If sts = 0 Then
    isMil = True
ElseIf sts = 1 Then
    isComm = True
ElseIf sts = 2 Then
    isother = True
End If

ICAOToCount = icao

    Select Case icao
    
        Case "B732", "B733", "B734", "B735", "B736", "B737", "B738"
            If isMil And icao = "B737" Then
                ICAOToCount = "C40"
            Else
                ICAOToCount = "B737"
            End If
            
        Case "B741", "B742", "B743", "B744", "B748"
            If isComm Then
                ICAOToCount = "B747"
            ElseIf icao = "B742" And isMil Then
                ICAOToCount = "E4"
            End If
            
        Case "B752"
            If isMil Then ICAOToCount = "C32" Else ICAOToCount = "B757"
            
        Case "B753"
            ICAOToCount = "B757"
            
        Case "B762", "B763"
            ICAOToCount = IIf(isother, "KC767", "B767")
            
        Case "BE20", "B190", "B350"
            If isMil Then ICAOToCount = "C12"
            
        Case "GLF3", "GLF4"
            If isother Then
                ICAOToCount = "U4"
            ElseIf isMil Then
                ICAOToCount = "C20"
            End If
            
        Case "C560"
            If isMil Then ICAOToCount = "UC35"
            
        Case "GLF5"
            If isMil Then ICAOToCount = "C37"
            
        Case "D328"
            ICAOToCount = "C146"
            
        Case "C30J"
            ICAOToCount = "C130"
            
        Case "DC10"
            If isMil Then ICAOToCount = "KC10"
        
        Case "K35R"
            ICAOToCount = "KC135"
        
        Case "R135"
            ICAOToCount = "RC135"
            
        Case "C135"
            ICAOToCount = "OC135"
            
        Case "E3TF"
            ICAOToCount = "E3"
        
        Case "B703"
            If isMil Then ICAOToCount = "E8"
        
        Case "B212"
            If isMil Then ICAOToCount = "UH1"
            
        Case "MV22", "CV22"
            ICAOToCount = "V22"
        
    End Select
    
End Function

Function fieldExists(ByVal sField As String, ByVal sTable As String) As Boolean
    err.clear
    fieldExists = False
    On Error GoTo setfalse
    If (DCount(sField, sTable) = 0) And err Then fieldExists = False Else fieldExists = True

setfalse:
End Function

Function userInitials() As String
    userInitials = Nz(DLookup("opInitials", "tblUserAuth", "username = '" & Environ$("username") & "'"))
End Function

Function LToZ(ByVal lcl As String) As Date
    Dim timezone As Integer
    timezone = DLookup("Timezone", "settings", "ID=1")
    If lcl = "" Then Exit Function
    
    LToZ = DateAdd("h", -timezone, lcl)
End Function

Function ZToL(ByVal zulu As String, Optional isTime As Boolean) As String
    Dim timezone As Integer
    timezone = DLookup("Timezone", "settings", "ID=1")
    If zulu = "" Then Exit Function
    
    ZToL = DateAdd("h", timezone, zulu)
    If isTime Then ZToL = Format(ZToL, "hh:nn")
End Function

