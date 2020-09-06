VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBarLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' By Adam Waller
' Last Modified:  12/16/05

'Private Const sngOffset As Single = 1.5    ' For Excel
Private Const sngOffset As Single = 15      ' For Access

Private mdblMax As Double   ' max value of progress bar
Private mdblVal As Double   ' current value of progress bar
Private mdblFullWidth As Double ' width of front label at 100%
Private mdblIncSize As Double
Private mblnHideCap As Boolean  ' display percent complete
Private mobjParent As Object    ' parent of back label
Private mlblBack As Access.Label     ' existing label for back
Private mlblFront As Access.Label   ' label created for front
Private mlblCaption As Access.Label ' progress bar caption
Private mdteLastUpdate As Date      ' Time last updated
Private mblnNotSmooth As Boolean    ' Display smooth bar by doevents after every update.

' This class displays a progress bar created
' from 3 labels.
' to use, just add a label to your form,
' and use this back label to position the
' progress bar.

Public Sub Initialize(BackLabel As Access.Label, FrontLabel As Access.Label, CaptionLabel As Access.Label)

    On Error GoTo 0    ' Debug Mode


    Dim objParent As Object ' could be a form or tab control
    Dim frm As Form

    Set mobjParent = BackLabel.Parent
    ' set private variables
    Set mlblBack = BackLabel
    Set mlblFront = FrontLabel
    Set mlblCaption = CaptionLabel

    ' set properties for back label
    With mlblBack
        .Visible = True
        .SpecialEffect = 2  ' sunken. Seems to lose when not visible.
    End With

    ' set properties for front label
    With mlblFront
        mdblFullWidth = mlblBack.width - (sngOffset * 2)
        .Left = mlblBack.Left + sngOffset
        .Top = mlblBack.Top + sngOffset
        .width = 0
        .Height = mlblBack.Height - (sngOffset * 2)
        .Caption = ""
        .BackColor = 8388608
        .BackStyle = 1
        .Visible = True
    End With

    ' set properties for caption label
    With mlblCaption
        .Left = mlblBack.Left + 2
        .Top = mlblBack.Top + 2
        .width = mlblBack.width - 4
        .Height = mlblBack.Height - 4
        .TextAlign = 2 'fmTextAlignCenter
        .BackStyle = 0 'fmBackStyleTransparent
        .Caption = "0%"
        .Visible = Not Me.HideCaption
        .ForeColor = 16777215   ' white
    End With
    'Stop

    Exit Sub

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Initialize", Erl
            Resume Next ' Resume at next line.
    End Select

End Sub

Private Sub Class_Terminate()

    On Error GoTo 0    ' Debug Mode

    On Error Resume Next
    mlblFront.Visible = False
    mlblCaption.Visible = False
    On Error GoTo 0    ' Debug Mode

    Exit Sub

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Class_Terminate", Erl
            Resume Next ' Resume at next line.
    End Select

End Sub

Public Property Get Max() As Double

    On Error GoTo 0    ' Debug Mode

    Max = mdblMax

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Max", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Let Max(ByVal dblMax As Double)

    On Error GoTo 0    ' Debug Mode

    mdblMax = dblMax

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Max", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Get Value() As Double

    On Error GoTo 0    ' Debug Mode

    Value = mdblVal

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Value", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Let Value(ByVal dblVal As Double)

    On Error GoTo 0    ' Debug Mode

    'update only if change is => 1%
    If (CInt(dblVal * (100 / mdblMax))) > (CInt(mdblVal * (100 / mdblMax))) Then
        mdblVal = dblVal
        update
    Else
        mdblVal = dblVal
    End If

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Value", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Get IncrementSize() As Double

    On Error GoTo 0    ' Debug Mode

    IncrementSize = mdblIncSize

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "IncrementSize", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Let IncrementSize(ByVal dblSize As Double)

    On Error GoTo 0    ' Debug Mode

    mdblIncSize = dblSize

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "IncrementSize", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Get HideCaption() As Boolean

    On Error GoTo 0    ' Debug Mode

    HideCaption = mblnHideCap

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "HideCaption", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Public Property Let HideCaption(ByVal blnHide As Boolean)

    On Error GoTo 0    ' Debug Mode

    mblnHideCap = blnHide

    Exit Property

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "HideCaption", Erl
            Resume Next ' Resume at next line.
    End Select

End Property

Private Sub update()

    On Error GoTo 0    ' Debug Mode

    Dim intPercent As Integer
    Dim dblWidth As Double
    'On Error Resume Next
    intPercent = mdblVal * (100 / mdblMax)
    dblWidth = mdblVal * (mdblFullWidth / mdblMax)
    mlblFront.width = dblWidth
    mlblCaption.Caption = intPercent & "%"
    'mlblFront.Parent.Repaint    ' may not be needed

    ' Use white or black, depending on progress
    If Me.Value > (Me.Max / 2) Then
        mlblCaption.ForeColor = 16777215   ' white
    Else
        mlblCaption.ForeColor = 0  ' black
    End If

    If mblnNotSmooth Then
        If mdteLastUpdate <> Now Then
            ' update every second.
            DoEvents
            mdteLastUpdate = Now
        End If
    Else
        DoEvents
    End If

    Exit Sub

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Update", Erl
            Resume Next ' Resume at next line.
    End Select

End Sub

Public Sub Increment()

    On Error GoTo 0    ' Debug Mode

    Dim dblVal As Double
    dblVal = Me.Value
    If dblVal < Me.Max Then
        Me.Value = dblVal + 1
        'Call Update
    End If

    Exit Sub

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Increment", Erl
            Resume Next ' Resume at next line.
    End Select

End Sub

Public Sub clear()

    On Error GoTo 0    ' Debug Mode

    Call Class_Terminate

    Exit Sub

ErrHandler:

    Select Case err.Number
        Case Else
            LogErr err, "clsLblProg", "Clear", Erl
            Resume Next ' Resume at next line.
    End Select

End Sub

Private Function ParentForm(ctlControl As Control) As String

    ' returns the name of the parent form
    Dim objParent As Object

    Set objParent = ctlControl

    Do While Not TypeOf objParent Is Form
       Set objParent = objParent.Parent
    Loop

    ' Now we should have the parent form
    ParentForm = objParent.Name

End Function

Public Property Get Smooth() As Boolean
    ' Display the progress bar smoothly.
    ' True by default, this property allows the call
    ' to doevents after every increment.
    ' If False, it will only update once per second.
    ' (This may increase speed for fast progresses.)
    '
    ' negative to set default to true
    Smooth = mblnNotSmooth
End Property

Public Property Let Smooth(ByVal IsSmooth As Boolean)
    mblnNotSmooth = Not IsSmooth
End Property

Private Sub LogErr(objErr, strMod, strProc, intLine)
    ' For future use.
End Sub
