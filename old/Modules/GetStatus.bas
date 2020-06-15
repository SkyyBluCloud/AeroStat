Attribute VB_Name = "GetStatus"
Option Compare Database

Function GetStatus(ByVal sts As Integer) As String
    Select Case sts
        Case 0
            GetStatus = ""
        Case 1
            GetStatus = "Pending"
        Case 2
            GetStatus = "Enroute"
        Case 3
            GetStatus = "Closed"
        Case 4
            GetStatus = "Cancelled"
    End Select
End Function

Function MarkCancelled(ByVal ctls As Controls)
    For Each ctl In ctls
        If UCase(TypeName(ctl)) = "TEXTBOX" Then
            ctl.BackColor = "#CF7B79"
        End If
    Next
End Function
