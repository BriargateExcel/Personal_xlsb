Attribute VB_Name = "CodeImport"
Option Explicit

Private Const Module_Name As String = "CodeImport."

Public Sub Import()
' todo: update comments
    '// Imports textual data from the file system such as VBA code to build the
    '// current active VBProject as specified in its configuration file.
    '// * Each code module file listed in the configuration file is imported into
    '//   the VBProject.
    '//   TODO Warn the user of overwrites
    '//   Modules with the same name are overwritten.
    '// * All references declared in the configuration file are
    '//   loaded into the project.
    '// * The project name is set to the project name specified by
    '//   the configuration file.
    ' Version 1.0
    ' Modified Import to use GetConfigFile
    
    Const RoutineName As String = Module_Name & "Import"
    On Error GoTo ErrorHandler
    
    Dim ProjectName As String
    
    SetTitle "Which Project and Group do you want to import?"
    
    GetModulesOfInterest
    
    If GetFormCanceled Then GoTo Done
    
    If WorkBookName = ThisWorkbook.Name Then
        MsgBox "Can not import into the workbook with the import code", _
               vbOKOnly Or vbCritical, _
               "Can Not Import into ThisWorkbook"
        GoTo Done
    End If

    '// Import code from listed module files
    Dim I As Long
    Dim FP As String
    Dim VBC As VBComponent
    
    For I = LBound(ModuleGroupArray, 1) To UBound(ModuleGroupArray, 1)
        Set VBC = VBAProject.VBComponents(ModuleGroupArray(I))
        FP = FolderPath & Application.PathSeparator & _
                    ModuleGroupArray(I) & FileExtension(VBC)
                    
        ImportModule VBAProject, ModuleGroupArray(I), FP
    Next I

    ' Set the references
    Dim Entry As Variant
    Dim Ref As VBAReferences_Table
    
    If Not (ModuleGroup = "Common" Or ModuleGroup = "Built") Then
        For Each Entry In ReferenceDict
            Set Ref = ReferenceDict(Entry)

            If Not CheckNameInCollection(Ref.Name, VBAProject.References) Then
                VBAProject.References.AddFromGuid _
                    GUID:=Ref.GUID, _
                    Major:=Ref.Major, _
                    Minor:=Ref.Minor
            End If
        Next Entry
    End If

Done:
    MsgBox "All modules successfully imported", _
           vbOKOnly Or vbInformation, _
           "Modules Imported Successfully"
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
End Sub ' Import

Private Sub ImportModule( _
        ByVal ThisProject As VBProject, _
        ByVal ModuleName As String, _
        ByVal ModulePath As String)
        
    '// Import a VBA code module
    
    Const RoutineName As String = Module_Name & "ImportModule"
    On Error GoTo ErrorHandler
    
    Dim ErrorNumber As Long
    
    Dim NameToCheck As String
    On Error Resume Next
    NameToCheck = ThisProject.VBComponents.Item(ModuleName).Name
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
    If ErrorNumber = 0 Then
        With VBAProject.VBComponents
            If .Item(ModuleName).Type <> vbext_ct_Document Then
                ' Can't remove a worksheet
                On Error Resume Next
                .Item(ModuleName).Name = ModuleName & "OLD"
                ErrorNumber = Err.Number
                If ErrorNumber = 0 Then
                    .Remove .Item(ModuleName & "OLD")
                Else
                    .Remove .Item(ModuleName)
                End If
                On Error GoTo ErrorHandler
                DoEvents
            End If
        End With
    End If
    
    On Error Resume Next
    
    Dim ComponentModule As VBComponent
    Set ComponentModule = ThisProject.VBComponents.Import(ModulePath)
    
    ErrorNumber = Err.Number
    If ErrorNumber = 60061 Then GoTo Done ' Module already in use
    On Error GoTo ErrorHandler
    
    Dim ExistingComponent As VBComponent
    Dim CodeMod As CodeModule
    Dim CodePasteMod As CodeModule
    
    If CheckNameInCollection(ModuleName, ThisProject.VBComponents) Then
        Set ExistingComponent = ThisProject.VBComponents(ModuleName)
        If ExistingComponent.Type = vbext_ct_Document Then
            
            Set CodePasteMod = ExistingComponent.CodeModule
            CodePasteMod.DeleteLines 1, CodePasteMod.CountOfLines
            
            Set CodeMod = ComponentModule.CodeModule
            
            If CodeMod.CountOfLines > 0 Then
                CodePasteMod.AddFromString CodeMod.Lines(1, CodeMod.CountOfLines)
            End If
            
            ThisProject.VBComponents.Remove ComponentModule
        End If
    End If
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' ImportModule

