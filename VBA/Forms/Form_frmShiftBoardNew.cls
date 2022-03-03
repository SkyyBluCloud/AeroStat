VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftBoardNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private varShiftID As Integer

Public Property Get Shift() As Integer
    Shift = varShiftID
End Property

Public Property Let Shift(ID As Integer)
    varShiftID = ID
    frmShiftDetailsNew.Requery
    AF3616.Requery
End Property
