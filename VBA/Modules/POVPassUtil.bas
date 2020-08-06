Attribute VB_Name = "POVPassUtil"
Option Compare Database

Public Function getPassNumber(ByVal newType As String) As String
    getPassNumber = newType & "-" & Format(Now, "yy") & "-" & _
        Format(Nz(DMax("right([Pass Number],3)", "qPOV", "[Pass Number] Is Not Null AND left([Pass Number],1) = '" & newType & "' AND mid([Pass Number],3,2) = format(now,'yy')"), 0) + 1, "000")
End Function

