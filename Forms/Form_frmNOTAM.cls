VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNOTAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub afi11208_Click()
DoCmd.OpenForm "regViewer", , , , , , "https://static.e-publishing.af.mil/production/1/af_a3/publication/afi11-208/afi11-208.pdf"
End Sub
