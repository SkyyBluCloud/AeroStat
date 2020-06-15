Attribute VB_Name = "sdas"
Option Compare Database

Public Function getuser()
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("tblUserAuth")
With rs
Do While Not .EOF
    Debug.Print !lastName & " " & !username
    .MoveNext
Loop
End With
End Function
