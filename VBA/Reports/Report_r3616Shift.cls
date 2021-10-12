VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_r3616Shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim iPageCount() As Integer
Public lclPage As Integer

Public Property Get pageCount(page) As Integer
    pageCount = iPageCount(page)
End Property

Private Sub btnEdit_Click()
DoCmd.OpenForm "frm3616", , , , , acDialog, shiftID
Parent.Requery
End Sub

Private Sub Detail_Print(cancel As Integer, PrintCount As Integer)
Dim eol As Double
eol = Me.Height
If lclPage = 0 Then lclPage = 1

    If lclPage <> PrintCount - 1 Then
        ReDim Preserve iPageCount(1 To PrintCount) As Integer
        lclPage = lclPage + 1
    End If
    
    'log "PrintCount = " & PrintCount, "r3616Shift.Print"
    
    iPageCount(PrintCount) = iPageCount(PrintCount) + 1

Me.Line (Me.Line43.Left, 0)-(Me.Line43.Left, eol)
'Me.Line (Me.Line44.Left, 0)-(Me.Line44.Left, eol)
'Me.Line (Me.Line92.Left, 0)-(Me.Line92.Left, eol)

End Sub
