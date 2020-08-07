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

'Private Declare ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
'                    ByVal hwnd As Long, _
'                    ByVal lpOperation As String, _
'                    ByVal lpFile As String, _
'                    ByVal lpParameters As String, _
'                    ByVal lpDirectory As String, _
'                    ByVal nShowCmd As Long) As Long

Private Sub Detail_Print(cancel As Integer, PrintCount As Integer)
Dim eol As Double
eol = Me.Height

Me.Line (Me.Line43.Left, 0)-(Me.Line43.Left, eol)
'Me.Line (Me.Line44.Left, 0)-(Me.Line44.Left, eol)
'Me.Line (Me.Line92.Left, 0)-(Me.Line92.Left, eol)

End Sub

'Private Sub savePDF_Click()
''Create nonexisting directories
'Dim f As String
'    DoCmd.SetWarnings False
'    f = DLookup("drivePrefix", "tblSettings") & "1. Events Logs\" & UCase(Format(rDate, "yyyy\\mm mmm yyyy\\d mmm yy ") & DLookup("shiftTitle", "tblShiftManager", Me.Filter) & " db") & ".pdf"
'    DoCmd.OutputTo acOutputReport, "r3616", acFormatPDF, f
'    MsgBox "To sign, go to: " & vbCrLf & "Tools -> Certificates -> Digitally Sign" & vbCrLf & vbCrLf & "Then draw a signature block under 'Signature'" & vbCrLf & "(This log is not valid without a signature)", vbInformation, "AeroStat"
'    ShellExecute Me.hwnd, "Open", f, vbNullString, "\\", 1
'    DoCmd.SetWarnings True
'End Sub

