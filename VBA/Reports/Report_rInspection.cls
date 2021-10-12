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
Option Explicit
Const maxColumnItems = 17

Private Sub Report_Load()
On Error GoTo errtrap
Dim start As Integer
Dim itemCount As Integer
Dim colCount As Integer
Dim colCats As Integer


    start = 1
    colCount = 0
    
    Dim col As Integer: For col = 1 To 3
    
        Dim cat: For cat = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
        
            'colCount = colCount + 1
            itemCount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & cat) + (cat * 2)
            
            colCats = cat
            'If there are too many items, exclude the last [cat], and move to the next [col]
            If itemCount > maxColumnItems Then
                colCats = cat - 1
                Exit For
            End If
        Next cat
        
        Me.Controls("inspectionItems" & col).Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & colCats
        start = colCats + 1
        colCount = 0
    Next col
    

'For i = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
'    colCount = colCount + 1
'    itemCount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & i) + (colCount * 2)
'    If itemCount >= maxColumnItems Then
'        colCats = IIf(itemCount > maxColumnItems, i - 1, i)
'        Exit For
'    End If
'Next
'inspectionItems2.Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & colCats
'
'start = colCats + 1
'colCount = 0
'For i = start To DCount("ID", "tblInspectionItemCategory", "ID >= 1")
'    colCount = colCount + 1
'    itemCount = DCount("categoryID", "tblInspectionItems", "categoryID Between " & start & " And " & i) + (colCount * 2)
'    If itemCount >= maxColumnItems Then
'        colCats = IIf(itemCount > maxColumnItems, i - 1, i)
'        Exit For
'    End If
'Next
'inspectionItems3.Report.RecordSource = "SELECT * FROM tblInspectionItemCategory INNER JOIN tblInspectionItems ON tblInspectionItemCategory.ID = tblInspectionItems.categoryID WHERE tblInspectionItemCategory.ID Between " & start & " And " & colCats

sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Name
    Resume Next
End Sub
