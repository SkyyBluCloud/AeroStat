Attribute VB_Name = "ShiftUtil"
Option Compare Database

Public Function isShiftClosed(ByVal shiftID As Integer) As Boolean

    If DLookup("closed", "tblShiftManager", "shiftid = " & shiftID) Then
        If Util.getOpInitials <> DLookup("right(superlead,2)", "tblshiftmanager", "shiftid = " & shiftID) Then
            MsgBox "This shift is closed. Only the AMOS can make changes.", vbInformation, "Checklist"
            Exit Function
        End If
    End If
    
    isShiftClosed = True
End Function
