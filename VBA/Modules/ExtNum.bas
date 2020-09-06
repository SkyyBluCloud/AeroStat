Attribute VB_Name = "ExtNum"
Option Compare Database

Function GetNum(ByVal pStr As String) As String

    Dim i As Integer, c As String
    If IsNull(pStr) Then GetNum = Null: Exit Function
    For i = 1 To Len(pStr)
        c = Mid(pStr, i, 1)
        If c Like "#" Then
            GetNum = GetNum & c
        ElseIf Len(GetNum) Then
            Exit Function
        End If
    Next i
End Function

Function DelNum(ByVal pStr As String) As String
    Dim num As Integer
    num = GetNum(pStr)
    DelNum = Replace(pStr, num, "")
End Function
