VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftDetailsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Current()
On Error GoTo errtrap
Dim amocs As New Collection
    
    Dim a: For Each a In Split(amoc, ";")
        amocs.add DLookup("ID", "tblUserAuth", "opinitials = '" & Right(a, 2) & "'")
    Next a
    
    With lstAmoc
    Dim i: For Each i In amocs
        Dim iSel: For iSel = 0 To .ListCount - 1
            If CInt(.Column(0, iSel)) = i Then
                .Selected(iSel) = True
                Exit For
            End If
        Next iSel
        
    Next i
    End With
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & "Form_Current"
End Sub

