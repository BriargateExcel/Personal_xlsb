Attribute VB_Name = "CodeBuildTables"
Option Explicit

Private Const Module_Name As String = "CodeBuildTables."

Public Sub MakeConfigurationTables()
    
    ' Builds the configuration tables
    
    Const RoutineName As String = Module_Name & "MakeConfigurationTables"
    On Error GoTo ErrorHandler
    
    SetTitle "Build tables for which Project?"
    
    GetModulesOfInterest
        
    If GetFormCanceled Then GoTo Done2
    
    If ModuleGroup = "Common" Then
        MsgBox "Can not build tables for Common Routines." & vbCrLf & _
            "Must be built by hand" & vbCrLf, _
            vbOKOnly, _
            "Can Not Build Tables for Common Routines"
        GoTo Done2
    End If
    
    If ModuleGroup = "Built" Then
        MsgBox "Can not build tables for TableBuilder built Routines." & vbCrLf & _
            "Must be built by hand" & vbCrLf, _
            vbOKOnly, _
            "Can Not Build Tables for TableBuilder built Routines"
        GoTo Done2
    End If
    
    ' Get the folders in which to store the code
    GetFolder "All"
    GetFolder "Built"
    GetFolder "Common"
    
    ' todo: build the worksheet and tables if missing

    '// Generate entries for modules not yet listed
    GetModules

    '// Generate entries for references in the current VBProject
    GetReferences
    
    '// Write changes to tables
    Dim ModuleList As VBAModuleList_Table
    Set ModuleList = New VBAModuleList_Table
    If Table.TryCopyDictionaryToTable(ModuleList, ModuleGroupDict, ModuleTable, , , True) Then
    Else
        ReportError "Error copying Module List to table", "Routine", RoutineName
        GoTo Done
    End If
    ' Module table has been updated
    
    Dim SourceFolder As VBASourceFolder_Table
    Set SourceFolder = New VBASourceFolder_Table
    If Table.TryCopyDictionaryToTable(SourceFolder, FolderPathDict, FolderPathTable, , , True) Then
    Else
        ReportError "Error loading Source Path", "Routine", RoutineName
        GoTo Done
    End If
    ' Paths table has been updated
    
    Dim RefList As VBAReferences_Table
    Set RefList = New VBAReferences_Table
    If Not ReferenceDict Is Nothing Then
        If Table.TryCopyDictionaryToTable(RefList, ReferenceDict, ReferenceTable, , , True) Then
        Else
            ReportError "Error loading References List", "Routine", RoutineName
            GoTo Done
        End If
    Else
        ClearTable ReferenceTable
    End If
    ' References table has been updated
    ' All tables have been updated
    
Done:
    MsgBox "Configuration Tables Built", vbOKOnly, "Configuration Tables Built"
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
End Sub ' MakeConfigurationTables

Private Sub GetFolder(ByVal FolderDescription As String)

    ' Used for lower level routines
    
    Const RoutineName As String = Module_Name & "GetFolder"
    On Error GoTo ErrorHandler
    
    If FolderPathDict.Exists(FolderDescription) Then
        FolderPathDict(FolderDescription).Path = _
            GetUserBasePath(FolderPathDict(FolderDescription).Path, _
                "Base path for " & FolderDescription & " modules")
    Else
    End If
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' GetFolder

Private Function GetUserBasePath( _
    ByVal InitialDirectory As String, _
    ByVal Message As String _
    ) As String
    
    ' Open the file dialog and capture the folder's path
    ' Version 1.0
    ' Part of refactoring GetBasePath
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "GetUserBasePath"
    On Error GoTo ErrorHandler
    
    Dim Response As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = Message
        .InitialFileName = InitialDirectory
        
        Response = .Show
        
        If Response <> 0 Then
            GetUserBasePath = .SelectedItems(1)
        Else
            GetUserBasePath = "No base path selected"
        End If
    End With
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetUserBasePath

Private Sub GetModules()

    ' This routine gathers a list of the modules in this project
    ' Compares that list with the existing config file
    ' Modifies the list of modules if the user desires
    
    Const RoutineName As String = Module_Name & "GetModules"
    On Error GoTo ErrorHandler
    
    Dim AddList As Collection
    Set AddList = New Collection
    
    Dim AddListStr As String
    AddListStr = vbNullString
    
    Dim ComponentModule As VBComponent
    Dim CreateNewEntry  As Boolean
    
    For Each ComponentModule In VBAProject.VBComponents
        CreateNewEntry = _
            ExportableModule(ComponentModule) And _
            Not ModuleGroupDict.Exists(ComponentModule.Name)

        If CreateNewEntry Then
            AddList.Add ComponentModule.Name
            AddListStr = AddListStr & ComponentModule.Name & vbNewLine
        End If
    Next ComponentModule

    ' Ask the user if they want to add new modules to the config file
    Dim UserResponse As Long
    Dim NewMod As VBAModuleList_Table
    Dim ModuleName As Variant
    
    If AddList.Count > 0 Then
        UserResponse = MsgBox( _
            Prompt:= _
                "There are some modules not listed in the configuration file which " & _
                "exist in the current project. Would you like to " & _
                "add these modules to the configuration file?" & _
                vbNewLine & _
                "Note: All modules are listed if there is no existing configuration file" & _
                vbNewLine & _
                "New modules:" & vbNewLine & _
                AddListStr, _
            Buttons:=vbYesNo + vbDefaultButton2, _
            Title:="New Modules")

        If UserResponse = vbYes Then
            For Each ModuleName In AddList
                Set NewMod = New VBAModuleList_Table
                If ModuleGroupDict.Exists(ModuleName) Then
                    ReportWarning "Duplicate module name", "Routine", RoutineName, "Module Name", ModuleName
                Else
                    NewMod.Module = ModuleName
                    ModuleGroupDict.Add ModuleName, NewMod
                End If
            Next ModuleName
        End If
    End If
    
    '// Ask user if they want to delete entries for missing modules
    ' Create the list of modules to potentially delete
    Dim DeleteList As Collection
    Set DeleteList = New Collection

    Dim DeleteListStr As String
    DeleteListStr = vbNullString

    Dim DeleteModule As Boolean
    
    For Each ModuleName In ModuleGroupDict
        DeleteModule = True

        If CheckNameInCollection(ModuleName, VBAProject.VBComponents) Then
            If ExportableModule(VBAProject.VBComponents(ModuleName)) Then
                DeleteModule = False
            End If
        End If

        If DeleteModule Then
            DeleteList.Add ModuleName
            DeleteListStr = DeleteListStr & ModuleName & vbNewLine
        End If
    Next ModuleName
    ' Now have a list of modules to potentially delete

    ' Ask the user if they want to delete any modules
    If DeleteList.Count > 0 Then
        UserResponse = MsgBox( _
                          Prompt:= _
                          "There are some modules listed in the configuration file which " & _
                          "haven't been found in the current project. Would you like to " & _
                          "remove these modules from the configuration file?" & _
                          vbNewLine & _
                          vbNewLine & _
                          "Missing modules:" & vbNewLine & _
                          DeleteListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="Missing Modules")

        If UserResponse = vbYes Then
            For Each ModuleName In DeleteList
                ModuleGroupDict.Remove ModuleName
            Next ModuleName
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
End Sub                                          ' GetModules

Private Sub GetReferences()
    ' This routine gathers a list of the references in this project
    ' Compares that list with the existing config file
    ' Modifies the list of reference if the user desires
    
    Const RoutineName As String = Module_Name & "GetReferences"
    On Error GoTo ErrorHandler
    
    Dim AddList As Collection
    Set AddList = New Collection
    
    Dim AddListStr As String
    AddListStr = vbNullString
    
    Dim Ref As Variant
    
    For Each Ref In VBAProject.References
        If Not Ref.BuiltIn Then
            If ReferenceToAdd(Ref) Then
                If ReferenceDict Is Nothing Then Set ReferenceDict = New Dictionary
                If Not ReferenceDict.Exists(Ref.Name) Then
                    AddList.Add Ref
                    AddListStr = AddListStr & Ref.Name & vbNewLine
                End If
            End If
        End If
    Next Ref

    ' Ask the user if they want to add new references to the config file
    Dim UserResponse As Long
    Dim NewRef As VBAReferences_Table
    
    If AddList.Count > 0 Then
        UserResponse = MsgBox( _
                          Prompt:= _
                          "There are some references not listed in the configuration file which " & _
                          "exist in the current project. Would you like to " & _
                          "add these references to the configuration file?" & _
                          vbNewLine & _
                          "Note: if the configuration file doesn't already exist, this will be a list of all references" & _
                          vbNewLine & _
                          "New references:" & vbNewLine & _
                          AddListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="New References")

        If UserResponse = vbYes Then
            Dim I As Long
            I = 1
            For Each Ref In AddList
                If ReferenceDict.Exists(Ref.Name) Then
                    ReportWarning "Duplicate reference name", "Routine", RoutineName, "Reference Name", Ref
                Else
                    Set NewRef = New VBAReferences_Table

                    NewRef.Name = Ref.Name
                    NewRef.Description = Ref.Description
                    NewRef.GUID = Ref.GUID
                    NewRef.Major = Ref.Major
                    NewRef.Minor = Ref.Minor
                    ReferenceDict.Add Ref.Name, NewRef
                End If
            Next Ref
        End If
    End If
    
    '// Ask user if they want to delete entries for missing references
    ' Create the list of modules to potentially delete
    Dim collDeleteList As Collection
    Set collDeleteList = New Collection

    Dim strDeleteListStr As String
    strDeleteListStr = vbNullString
    
    If Not ReferenceDict Is Nothing Then
        For Each Ref In ReferenceDict
            If Not CheckNameInCollection(Ref, VBAProject.References) Then
                collDeleteList.Add Ref
                strDeleteListStr = strDeleteListStr & Ref & vbNewLine
            End If
        Next Ref
    
        ' Ask the user if they want to delete any references
        If collDeleteList.Count > 0 Then
            UserResponse = MsgBox( _
                              Prompt:="There are some references listed in the configuration file which " & _
                                       "haven't been found in the current project. Would you like to " & _
                                       "remove these references from the configuration file?" & vbNewLine & _
                                       vbNewLine & _
                                       "Missing references:" & vbNewLine & _
                                       strDeleteListStr, _
                              Buttons:=vbYesNo + vbDefaultButton2, _
                              Title:="Missing References")
    
            If UserResponse = vbYes Then
                For Each Ref In collDeleteList
                    ReferenceDict.Remove Ref
                Next Ref
            End If
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
End Sub                                          ' GetReferences

Private Function ExportableModule(ByVal ComponentModule As VBComponent) As Boolean
    '// Is the given module exportable by this tool?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ExportableModule"
    On Error GoTo ErrorHandler
    
    ExportableModule = _
        (Not ModuleEmpty(ComponentModule)) And (Not FileExtension(ComponentModule) = vbNullString)
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ExportableModule

Private Function ModuleEmpty(ByVal ThisModule As VBComponent) As Boolean
    '// Check if a code module is effectively empty.
    '// effectively empty should be functionally and semantically equivalent to
    '// actually empty.
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ModuleEmpty"
    On Error GoTo ErrorHandler
    
    ModuleEmpty = True

    Dim NumLines As Long
    NumLines = ThisModule.CodeModule.CountOfLines
    
    Dim LineNumber As Long
    Dim CurrentLine As String
    
    For LineNumber = 1 To NumLines
        CurrentLine = ThisModule.CodeModule.Lines(LineNumber, 1)
        
        If Not (CurrentLine = "Option Explicit" Or CurrentLine = vbNullString) Then
            ModuleEmpty = False
            GoTo Done
        End If
    Next LineNumber
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ModuleEmpty

Private Function ReferenceToAdd(ByVal ThisRef As Reference) As Boolean
    ' This routine determies if this reference needs to be saved
    ' Returns True if the reference must be saved
    ' Returns False if this reference does not need to be saved
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ReferenceToAdd"
    On Error GoTo ErrorHandler
    
    Dim NeededRef As Boolean
    Dim ErrorNumber As Long
    
    On Error Resume Next
    NeededRef = (ThisRef.Name = "stdole")
    ErrorNumber = Err.Number
    On Error GoTo 0
    If ErrorNumber <> 0 Then
        ReferenceToAdd = False ' Bad reference; skip
        GoTo Done
    End If
    
    Select Case ThisRef.Name
    Case "stdole"
        ReferenceToAdd = False
    Case "MSForms"
        ReferenceToAdd = False
    Case "Office"
        ReferenceToAdd = False
    Case Else
        ReferenceToAdd = True
    End Select
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' ReferenceToAdd

