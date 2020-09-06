Attribute VB_Name = "sdas"
Option Compare Database

Public Function getuser()
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("tblUserAuth")
With RS
Do While Not .EOF
    Debug.Print !lastName & " " & !username
    .MoveNext
Loop
End With
End Function
