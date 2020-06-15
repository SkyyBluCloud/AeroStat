Attribute VB_Name = "frmParkingMapUtil"
Option Compare Database

Public Function checkSyntax(ByVal Spot As String) As Boolean
If IsNull(Spot) Then Exit Function
Spot = Replace(Spot, " ", "")
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("tblParkingManagement")

'    For Each s In Split(Spot, ",")
'        If InStr(1, s, "-") > 0 Then
'
        
    
    
End Function
