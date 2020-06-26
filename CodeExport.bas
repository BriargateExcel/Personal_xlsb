Attribute VB_Name = "CodeExport"
Option Explicit

Private Const Module_Name As String = "CodeExport."

Public Sub Export()
' todo: redo comments

    '// Exports code modules and cleans the current active VBProject as specified
    '// by the project's configuration file.
    '// * Any code module in the VBProject which is listed in the configuration
    '//   file is exported to the configured path.
    '// * code modules which were exported are deleted or cleared.
    '// * References loaded in the Project which are listed in the configuration
    '//   file is deleted.
    
    Const RoutineName As String = Module_Name & "Export"
    On Error GoTo ErrorHandler
    
    SetTitle "Which Project and Group do you want to export?"
    
    GetModulesOfInterest
    
    If GetFormCanceled Then GoTo Done2
    
    '// Export all modules listed in the configuration
    Dim ComponentModule As VBComponent
    
    Dim I As Long
    
    For I = LBound(ModuleGroupArray, 1) To UBound(ModuleGroupArray, 1)
        ' TODO Provide a warning if module listed in configuration is not found
        If CheckNameInCollection(ModuleGroupArray(I), VBAProject.VBComponents) Then
            Set ComponentModule = VBAProject.VBComponents(ModuleGroupArray(I))
            
            Dim Dest As String
            Dest = FolderPath & Application.PathSeparator & ModuleGroupArray(I) & FileExtension(ComponentModule)
            ComponentModule.Export Dest
        End If
    Next I

Done:
    MsgBox "All modules successfully exported", _
           vbOKOnly Or vbInformation, _
           "Modules Exported Successfully"
    GoTo Done2
Halted:
    ' Use the Halted exit point after giving the user a message
    '   describing why processing did not run to completion
    MsgBox "Abnormal exit"
Done2:
    CloseErrorFile
    TurnOnAutomaticProcessing
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    TurnOnAutomaticProcessing
    CloseErrorFile
End Sub ' Export
