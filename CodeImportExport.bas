Attribute VB_Name = "CodeImportExport"
Option Explicit

' Originally from: https://github.com/spences10/VBA-IDE-Code-Export

Const Module_Name As String = "CodeImportExport."

Private Type CodeType
    FormCanceled As Boolean
    FormDeleted As Boolean
    
    ModuleList As Dictionary
    ModuleTable As ListObject
    
    PathFolder As Dictionary
    PathTable As ListObject
    Path As String
    
    ReferencesList As Dictionary
    ReferencesTable As ListObject
    
    Project As VBProject
    
    Workbook As Workbook
    Worksheet As Worksheet
    
    CopyCommonModules As Boolean
End Type

Private This As CodeType

Private Const VBAMakeFile As String = "VBA Make File"
Private Const VBACommonFile As String = "VBA Common Make File"

Private Const VBAModuleList As String = "VBAModuleList"
Private Const VBACommonModuleList As String = "VBACommonModuleList"

Private Const VBASourceFolder As String = "VBASourceFolder"
Private Const VBACommonSourceFolder As String = "VBACommonSourceFolder"

Private Const VBAReferences As String = "VBAReferences"

Public FSO As FileSystemObject

Public Sub MakeConfigFile()
    
    ' Builds the configuration tables
    
    Const RoutineName As String = Module_Name & "MakeConfigFile"
    On Error GoTo ErrorHandler
    
    PopulateTables "Build configuration tables for which project?"
    
    If This.FormCanceled Then Exit Sub
    
    If This.CopyCommonModules And This.Workbook.Name <> "PERSONAL.xlsb" Then
        MsgBox "Can not build tables for Common Routines." & vbCrLf & _
            "Must be built by hand" & vbCrLf & _
            "Go to View -> Window/Unhide to make PERSONAL.xlsb visible.", _
            vbOKOnly, _
            "Can Not Build Tables for Common Routines"
        Exit Sub
    End If
    
    ' Get the folder in which to store the code
    Dim BasePath As String
    BasePath = GetUserBasePath(This.Path)
    If BasePath = "No base path selected" Then
        MsgBox "No base path selected. No configuration file created.", _
               vbOKOnly Or vbCritical, _
               "No Base Path Selected"
        Exit Sub
    Else
        This.Path = BasePath
        This.PathFolder.Items(0).Path = BasePath
    End If
    ' Now have the source code folder
    
    '// Generate entries for modules not yet listed
    GetModules

    '// Generate entries for references in the current VBProject
    GetReferences
    
    '// Write changes to tables
    Dim Tbl As ListObject
    Dim SourceFolder As VBASourceFolder_Table
    Set SourceFolder = New VBASourceFolder_Table
    
    If Table.TryCopyDictionaryToTable(SourceFolder, This.PathFolder, This.PathTable, , , True) Then
    Else
        ReportError "Error loading Source Path", "Routine", RoutineName
        GoTo Done
    End If
    
    Dim RefList As VBAReferences_Table
    Set RefList = New VBAReferences_Table
    
    If Table.TryCopyDictionaryToTable(RefList, This.ReferencesList, This.ReferencesTable, , , True) Then
    Else
        ReportError "Error loading References List", "Routine", RoutineName
        GoTo Done
    End If
    Dim ModuleList As VBAModuleList_Table
    Set ModuleList = New VBAModuleList_Table
    
    If Table.TryCopyDictionaryToTable(ModuleList, This.ModuleList, This.ModuleTable, , , True) Then
    Else
        ReportError "Error copying Module List to table", "Routine", RoutineName
        GoTo Done
    End If
    ' Tables have been updated
    
    MsgBox "Configuration Tables Built", vbOKOnly, "Configuration Tables Built"

exitSub:
    Exit Sub
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
    CloseErrorFile
End Sub ' MakeConfigFile

Public Sub Export()

    '// Exports code modules and cleans the current active VBProject as specified
    '// by the project's configuration file.
    '// * Any code module in the VBProject which is listed in the configuration
    '//   file is exported to the configured path.
    '// * code modules which were exported are deleted or cleared.
    '// * References loaded in the Project which are listed in the configuration
    '//   file is deleted.
    
    Const RoutineName As String = Module_Name & "Export"
    On Error GoTo ErrorHandler
    
    PopulateTables "Export which project?"
    
    If This.FormCanceled Then Exit Sub
    
    EnsurePath This.Path
    
    '// Export all modules listed in the configuration
    Dim ModuleName As Variant
    Dim ComponentModule As VBComponent
    
    For Each ModuleName In This.ModuleList
        ' TODO Provide a warning if module listed in configuration is not found
        If CheckNameInCollection(ModuleName, This.Project.VBComponents) Then
            Set ComponentModule = This.Project.VBComponents(ModuleName)
            
            Dim Dest As String
            Dest = This.Path & Application.PathSeparator & ModuleName & FileExtension(ComponentModule)
            ComponentModule.Export Dest

            If This.FormDeleted Then
                If ComponentModule.Type = vbext_ct_Document Then
                    ComponentModule.CodeModule.DeleteLines 1, ComponentModule.CodeModule.CountOfLines
                Else
                    This.Project.VBComponents.Remove ComponentModule
                End If
            End If
        End If
    Next ModuleName

    '// Remove all references listed
    If Not This.CopyCommonModules Then
        If This.FormDeleted Then
            For Each ModuleName In This.ModuleList
                If CheckNameInCollection(ModuleName, This.Project.References) Then
                    This.Project.References.Remove This.Project.References(ModuleName)
                End If
            Next ModuleName
        End If
    End If

    MsgBox "All modules successfully exported", _
           vbOKOnly Or vbInformation, _
           "Modules Exported Successfully"
NoAction:
    Exit Sub
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
    CloseErrorFile
End Sub                                          ' Export

Public Sub Import()

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
    
    PopulateTables "Import which project?"
    
    If This.FormCanceled Then Exit Sub
    
    If This.Project.Name = "Personal" Then
        MsgBox "Can not import into PERSONAL.xlsb", _
               vbOKOnly Or vbCritical, _
               "Can Not Import PERSONAL.xlsb"
        Exit Sub
    End If

    '// Import code from listed module files
    Dim ModuleName As Variant
    Dim ComponentModule As VBComponent
    
    For Each ModuleName In This.ModuleList
        Set ComponentModule = This.Project.VBComponents(ModuleName)
        ImportModule _
            This.Project, _
            ModuleName, _
            This.Path & Application.PathSeparator & _
                This.ModuleList(ModuleName).Module & "." & This.ModuleList(ModuleName).Extension
            
    Next ModuleName

    '// Add references listed in the config file
    Dim Entry As Variant
    Dim Ref As VBAReferences_Table
    
    If Not This.CopyCommonModules Then
        For Each Entry In This.ReferencesList
            Set Ref = This.ReferencesList(Entry)
            
            If Not CheckNameInCollection(Ref.Name, This.Project.References) Then
                This.Project.References.AddFromGuid _
                    GUID:=Ref.GUID, _
                    Major:=Ref.Major, _
                    Minor:=Ref.Minor
            End If
        Next Entry
    End If
    
    MsgBox "All modules successfully imported", _
           vbOKOnly Or vbInformation, _
           "Modules Imported Successfully"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
    CloseErrorFile
End Sub                                          ' Import

Private Sub GetProject( _
    ByVal TitleText As String, _
    ByRef ThisProject As VBProject, _
    ByRef Wkbk As Workbook)
    
    Const RoutineName As String = Module_Name & "GetProject"
    On Error GoTo ErrorHandler
    
    GitForm.ProjectList.Clear
    
    Dim VBAProj As Variant
    For Each VBAProj In Application.VBE.VBProjects
        GitForm.ProjectList.AddItem VBAProj.Name
    Next VBAProj
    
    If GitForm.ProjectList.ListCount > 0 And _
       GitForm.ProjectList.List(0) = "Personal" _
       Then
        GitForm.ProjectList.Text = GitForm.ProjectList.List(1)
    Else
        GitForm.ProjectList.Text = GitForm.ProjectList.List(0)
    End If
    
    GitForm.Caption = TitleText
    GitForm.Show
    
    Dim SelectedProject As VBProject
    
    If This.FormCanceled Then
        ' Either Cancel button or dialog close button (red X) selected
        MsgBox TitleText & " Canceled by User", _
               vbOKOnly Or vbInformation, _
               "Cancel Selected"
        Set SelectedProject = Nothing
        Set Wkbk = Nothing
        Exit Sub
    Else
        Set SelectedProject = Application.VBE.VBProjects(GitForm.ProjectList.Value)
    End If

    If SelectedProject.Protection = 1 Then
        MsgBox "This project is protected, not possible to export the code"
        Exit Sub
    End If
    
    Set ThisProject = SelectedProject
    
    Dim Entry As Workbook
    For Each Entry In Workbooks
        If Entry.VBProject.Name = SelectedProject.Name Then
            Set Wkbk = Entry
            Exit For
        End If
    Next Entry
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' GetProject

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
    
    For Each ComponentModule In This.Project.VBComponents
        CreateNewEntry = _
                           ExportableModule(ComponentModule) And _
                           Not This.ModuleList.Exists(ComponentModule.Name)

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
                If This.ModuleList.Exists(ModuleName) Then
                    ReportWarning "Duplicate module name", "Routine", RoutineName, "Module Name", ModuleName
                Else
                    NewMod.Module = ModuleName
                    NewMod.Extension = FileExtension(This.Project.VBComponents(ModuleName))
                    This.ModuleList.Add ModuleName, NewMod
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
    
    For Each ModuleName In This.ModuleList
        DeleteModule = True

        If CheckNameInCollection(ModuleName, This.Project.VBComponents) Then
            If ExportableModule(This.Project.VBComponents(ModuleName)) Then
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
                This.ModuleList.Remove ModuleName
            Next ModuleName
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' GetModules

Private Function GetConfigDirectory(ByVal InitialDirectory As String) As String
    
    ' Open the file dialog and capture the folder's path
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "GetConfigDirectory"
    On Error GoTo ErrorHandler
    
    Dim Response As Long
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Choose Configuration File Directory"
        .InitialFileName = InitialDirectory
        
        Response = .Show
        
        If Response <> 0 Then
            GetConfigDirectory = .SelectedItems(1)
        Else
            GetConfigDirectory = "No directory selected"
        End If
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetConfigDirectory

Private Function GetUserBasePath(ByVal InitialDirectory As String) As String
    
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
        .Title = "Choose Base Path Directory"
        .InitialFileName = InitialDirectory
        
        Response = .Show
        
        If Response <> 0 Then
            GetUserBasePath = .SelectedItems(1)
        Else
            GetUserBasePath = "No base path selected"
        End If
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetUserBasePath

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
    
    For Each Ref In This.Project.References
        If Not Ref.BuiltIn Then
            If ReferenceToAdd(Ref) Then
                If Not This.ReferencesList.Exists(Ref.Name) Then
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
                If This.ReferencesList.Exists(Ref.Name) Then
                    ReportWarning "Duplicate reference name", "Routine", RoutineName, "Reference Name", Ref
                Else
                    Set NewRef = New VBAReferences_Table
                    
                    NewRef.Name = Ref.Name
                    NewRef.Description = Ref.Description
                    NewRef.GUID = Ref.GUID
                    NewRef.Major = Ref.Major
                    NewRef.Minor = Ref.Minor
                    This.ReferencesList.Add Ref.Name, NewRef
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
    
    For Each Ref In This.ReferencesList
        If Not CheckNameInCollection(Ref, This.Project.References) Then
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
                This.ReferencesList.Remove Ref
            Next Ref
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' GetReferences

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
        Exit Function
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
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ReferenceToAdd

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
    On Error GoTo 0
    If ErrorNumber = 0 Then
        Dim VBC As VBComponent
        Set VBC = ThisProject.VBComponents.Item(ModuleName)
        If ThisProject.VBComponents.Item(ModuleName).Type <> vbext_ct_Document Then
            ' Can't remove a worksheet
            ThisProject.VBComponents.Remove VBC
        End If
    Else
        Exit Sub
    End If
    
    On Error Resume Next
    
    Dim ComponentModule As VBComponent
    Set ComponentModule = ThisProject.VBComponents.Import(ModulePath)
    
    ErrorNumber = Err.Number
    If ErrorNumber = 60061 Then Exit Sub         ' Module already in use
    If ErrorNumber = 53 Then Exit Sub         ' Module does not exist
    On Error GoTo 0
    
    Dim NewComponent As VBComponent
    Dim ExistingComponent As VBComponent
    Dim CodeMod As CodeModule
    Dim CodePasteMod As CodeModule
    
    If ComponentModule.Name <> ModuleName Then
        If CheckNameInCollection(ModuleName, ThisProject.VBComponents) Then
            Set ExistingComponent = ThisProject.VBComponents(ModuleName)
            If ExistingComponent.Type = vbext_ct_Document Then
                
                Set CodePasteMod = ExistingComponent.CodeModule
                CodePasteMod.DeleteLines 1, CodePasteMod.CountOfLines
                
                Set CodeMod = NewComponent.CodeModule
                
                If CodeMod.CountOfLines > 0 Then
                    CodePasteMod.AddFromString CodeMod.Lines(1, CodeMod.CountOfLines)
                End If
                
                ThisProject.VBComponents.Remove NewComponent
            Else
                ExistingComponent.Name = ExistingComponent.Name & "_remove"
                ThisProject.VBComponents.Remove ExistingComponent
                NewComponent.Name = ModuleName   ' TODO fails on work computer
                ThisProject.VBComponents.Remove ExistingComponent
            End If
        Else
            NewComponent.Name = ModuleName
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ImportModule

Private Function ExportableModule(ByVal ComponentModule As VBComponent) As Boolean
    '// Is the given module exportable by this tool?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ExportableModule"
    On Error GoTo ErrorHandler
    
    ExportableModule = _
        (Not ModuleEmpty(ComponentModule)) And (Not FileExtension(ComponentModule) = vbNullString)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
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
            Exit Function
        End If
    Next LineNumber
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ModuleEmpty

Private Function FileExtension(ByVal ThisModule As VBComponent) As String
    '// The appropriate file extension for exporting the given module
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "FileExtension"
    On Error GoTo ErrorHandler
    
    Select Case ThisModule.Type
    Case vbext_ct_StdModule
        FileExtension = ".bas"
    Case vbext_ct_ClassModule, vbext_ct_Document
        FileExtension = ".cls"
    Case vbext_ct_MSForm
        FileExtension = ".frm"
    Case Else
        FileExtension = vbNullString
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' FileExtension

Private Sub EnsurePath(ByVal Path As String)
    '// Ensure path to a file exists. Creates missing folders.
    
    Const RoutineName As String = Module_Name & "EnsurePath"
    On Error GoTo ErrorHandler
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim ParentPath As String
    ParentPath = FSO.GetParentFolderName(Path)

    If ParentPath <> vbNullString Then
        EnsurePath ParentPath
        If Not FSO.FolderExists(ParentPath) Then
            If FSO.FileExists(ParentPath) Then
                ReportError "No path exists", _
                            "Path", ParentPath
            Else
                FSO.CreateFolder (ParentPath)
            End If
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' EnsurePath

Public Sub LetGitFormCanceled(ByVal Canx As Boolean)
    This.FormCanceled = Canx
End Sub ' LetGitFormCanceled

Public Sub CopyCommonModules(ByVal CopyCommon As Boolean)
    This.CopyCommonModules = CopyCommon
End Sub

Private Sub PopulateTables(ByVal Title As String)

    ' Reads all the data from the tables
    
    Const RoutineName As String = Module_Name & "PopulateTables"
    On Error GoTo ErrorHandler
    
    GetProject Title, This.Project, This.Workbook
    
    If This.FormCanceled Then Exit Sub
    
    Dim ModuleList As VBAModuleList_Table
    Dim SourceFolder As VBASourceFolder_Table
    Dim RefList As VBAReferences_Table
    Dim ErrorNumber As Long
    
    If This.CopyCommonModules Then
        On Error Resume Next
        Set This.Worksheet = This.Workbook.Worksheets(VBACommonFile)
        ErrorNumber = Err.Number
        If ErrorNumber <> 0 Then
            GoTo Done
        End If
        On Error GoTo ErrorHandler
        
        Set ModuleList = New VBAModuleList_Table
        
        Set This.ModuleTable = This.Worksheet.ListObjects(VBACommonModuleList)
        
        If Table.TryCopyTableToDictionary( _
            ModuleList, This.ModuleList, This.ModuleTable, False) _
        Then
        Else
            ReportError "Error loading Module List", "Routine", RoutineName
            GoTo Done
        End If
        
        Set SourceFolder = New VBASourceFolder_Table
        
        Set This.PathTable = This.Worksheet.ListObjects(VBACommonSourceFolder)
           
        If Table.TryCopyTableToDictionary( _
            SourceFolder, This.PathFolder, This.PathTable, False) _
        Then
            This.Path = This.PathFolder.Items(0).Path
        Else
            ReportError "Error loading Source Path", "Routine", RoutineName
            GoTo Done
        End If
    Else ' This.CopyCommonModules
        Set This.Worksheet = This.Workbook.Worksheets(VBAMakeFile)
        
        Set ModuleList = New VBAModuleList_Table
        
        Set This.ModuleTable = This.Worksheet.ListObjects(VBAModuleList)
        
        If Table.TryCopyTableToDictionary( _
            ModuleList, This.ModuleList, This.ModuleTable, False) _
        Then
        Else
            ReportError "Error loading Module List", "Routine", RoutineName
            GoTo Done
        End If
        
        Set SourceFolder = New VBASourceFolder_Table
        
        Set This.PathTable = This.Worksheet.ListObjects(VBASourceFolder)
           
        If Table.TryCopyTableToDictionary( _
            SourceFolder, This.PathFolder, This.PathTable, False) _
        Then
            This.Path = This.PathFolder.Items(0).Path
        Else
            ReportError "Error loading Source Path", "Routine", RoutineName
            GoTo Done
        End If
        
        Set RefList = New VBAReferences_Table
        
        Set This.ReferencesTable = This.Worksheet.ListObjects(VBAReferences)
        
        If Table.TryCopyTableToDictionary( _
            RefList, This.ReferencesList, This.ReferencesTable, False) _
        Then
        Else
            ReportError "Error loading References List", "Routine", RoutineName
            GoTo Done
        End If
    End If ' This.CopyCommonModules
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' PopulateTables

