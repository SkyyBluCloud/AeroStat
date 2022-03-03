VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDateSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Property Let date1(D1 As Variant)
    ctlDate1 = Nz(D1)
End Property

Public Property Get date1() As Variant
    date1 = ctlDate1
End Property

Public Property Let Category(cat As Variant)
    cboCategory = cat
End Property

Public Property Get Category() As Variant
    Category = cboCategory
End Property

Public Property Let search(s As String)
    
End Property
