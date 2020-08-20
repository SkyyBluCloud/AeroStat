VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Const maxColumnItems = 21

Private Sub Report_Load()
On Error GoTo errtrap
Dim start As Integer
Dim icount As Integer
Dim cCount As Integer
Dim loadCats As Integer


start = 1
cCount = 0
For i = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
    cCount = cCount + 1
    icount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & i) + (cCount * 2)
    If icount >= maxColumnItems Then
        loadCats = IIf(icount > maxColumnItems, i - 1, i)
        Exit For
    End If
Next
inspectionItems1.Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & loadCats

start = loadCats + 1
cCount = 0
For i = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
    cCount = cCount + 1
    icount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & i) + (cCount * 2)
    If icount >= maxColumnItems Then
        loadCats = IIf(icount > maxColumnItems, i - 1, i)
        Exit For
    End If
Next
inspectionItems2.Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & loadCats

start = loadCats + 1
cCount = 0
For i = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
    cCount = cCount + 1
    icount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & i) + (cCount * 2)
    If icount >= maxColumnItems Then
        loadCats = IIf(icount > maxColumnItems, i - 1, i)
        Exit For
    End If
Next
inspectionItems3.Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & loadCats

sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Name
    Resume Next
End Sub
