VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnAirfield_Click()
DoCmd.OpenForm "frmAirfieldEquipmentSurfaces"
End Sub

Private Sub btnBack_Click()
    Select Case tabCtl
    Case 0
        Exit Sub
        
    Case 2
        btnNext.Enabled = True
    
    End Select
    tabCtl = tabCtl - 1
End Sub

Private Sub btnBaseAcft_Click()
DoCmd.OpenForm "frmBaseAcft"
End Sub

Private Sub btnBrowse_Click()
Dim fd As Office.FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Filters.clear
        .Filters.add "Access Database", "*.accdb"
        .title = "Select AeroStat BACKEND file"
        
        If .show Then
            Dim s: For Each s In .SelectedItems
                txtBackend = s
            Next
        End If
    End With
    
    btnNext.Enabled = Util.relinkTables(txtBackend)
    
End Sub

Private Sub btnCreateBackend_Click()
On err GoTo errtrap
Dim fd As Office.FileDialog: Set fd = Access.FileDialog(msoFileDialogSaveAs)
Dim fso As New FileSystemObject
Dim f, fr, fc
Dim acc As New Access.Application
Dim saveLocation As String

    With fd 'File dialog for save location
        .title = "Save New Backend"
        
        If .show Then
            Dim s: For Each s In .SelectedItems
                saveLocation = s
            Next
            fso.DeleteFile saveLocation
        'Else
            'Cancelled by user
        End If
    End With
    
    With acc 'Create the new database file
        Dim schema As String
        
        .DBEngine.CreateDatabase saveLocation, dbLangGeneral
        .OpenCurrentDatabase saveLocation, True
        schema = .CurrentProject.Path & "\Schema\"
        
        If Not fso.FolderExists(schema) Then
            Set fd = Access.FileDialog(msoFileDialogFolderPicker)
            With fd
                .title = "Select schema folder"
                If .show Then
                    Dim s1: For Each s1 In .SelectedItems
                        schema = s1
                    Next
                Else
                    'Cancelled by user
                End If
            End With
                
        Else
            
        End If
        
        Set fr = fso.GetFolder(schema)
        Set fc = fr.Files
        For Each f In fc 'Lookup each file in schema folder, then import
            log "Creating table from XML: " & f.Name, "frmSetup.btnCreateBackend"
            .ImportXML f, acStructureOnly
        Next
        .CloseCurrentDatabase
    End With
    log "Done!", "frmSetup.btnCreateBackend"
    tabCtl = tabCtl + 1
    
sExit:
    'Cleanup
    Set acc = Nothing
    Set fd = Nothing
    Set fso = Nothing
    Set fr = Nothing
    Set fc = Nothing
    Exit Sub
errtrap:
    Select Case err
    Case 76
        MsgBox Error$ & " (" & err & ")", vbCritical, "Error"
    End Select
    Resume sExit
End Sub

Private Sub btnNext_Click()
On Error Resume Next
    Select Case tabCtl
    Case tabCtl.Pages.Count - 1 'Last page
        CurrentDb.Execute "UPDATE lclver SET firstrun = False"
        DoCmd.Close acForm, Me.Name
        
    Case 1 'License
        btnNext.Caption = "Next"
    
    Case 3 'Global Settings
        Dim fail As Boolean
        Dim ctl: For Each ctl In Me.frmSetupGlobalSettings.Form
            If ctl.Tag = "!" Then
                If Nz(ctl.Value) = "" Then
                    ctl.BorderColor = RGB(255, 0, 0)
                    fail = True
                End If
            End If
        Next
        If fail Then Exit Sub
        
    End Select
    tabCtl = tabCtl + 1
    
End Sub

Private Sub btnShiftDef_Click()
    DoCmd.OpenForm "frmShiftDef", , , , , acDialog
End Sub

Private Sub tabCtl_Change()

    Select Case tabCtl
    Case 1 'License
        btnNext.Caption = "Accept"
        
    Case 2 'Backend
        btnNext.Enabled = Nz(txtBackend) <> ""
    
    Case 3 'Global Settings
        
        
    Case tabCtl.Pages.Count - 1 'Last page
        btnNext.Caption = "Finish"
    
    End Select

End Sub
