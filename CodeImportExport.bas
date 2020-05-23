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
End Type

Private This As CodeType

Private Const VBAMakeFile As String = "VBA Make File"
Private Const VBAModuleList As String = "VBAModuleList"
Private Const VBASourceFolder As String = "VBASourceFolder"
Private Const VBAReferences As String = "VBAReferences"

Public FSO As FileSystemObject

Public Sub MakeConfigFile()
    
    ' Builds the configuration tables
    
    Const RoutineName As String = Module_Name & "MakeConfigFile"
    On Error GoTo ErrorHandler
    
    PopulateTables "Build configuration tables for which project?"
    
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
    GetReferences This.Project
    
    '// Write changes to tables
    Dim ModuleList As VBAModuleList_Table
    Set ModuleList = New VBAModuleList_Table
    
    Dim Wksht As Worksheet
    Set Wksht = This.Workbook.Worksheets("VBA Make File")
    
    Dim Tbl As ListObject
    Set Tbl = Wksht.ListObjects(VBAModuleList)
    
    If Table.TryCopyDictionaryToTable(ModuleList, This.ModuleList, Tbl, , , True) Then
    Else
        ReportError "Error copying Module List to table", "Routine", RoutineName
        GoTo Done
    End If
    
    Dim SourceFolder As VBASourceFolder_Table
    Set SourceFolder = New VBASourceFolder_Table
    Set Tbl = Wksht.ListObjects(VBASourceFolder)
    
    If Table.TryCopyDictionaryToTable(SourceFolder, This.PathFolder, Tbl, , , True) Then
    Else
        ReportError "Error loading Source Path", "Routine", RoutineName
        GoTo Done
    End If
    
    Dim RefList As VBAReferences_Table
    Set RefList = New VBAReferences_Table
    Set Tbl = Wksht.ListObjects(VBAReferences)
    
    If Table.TryCopyDictionaryToTable(RefList, This.ReferencesList, Tbl, , , True) Then
    Else
        ReportError "Error loading References List", "Routine", RoutineName
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
End Sub                                          ' MakeConfigFile

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
    
    EnsurePath This.Path
    
    '// Export all modules listed in the configuration
    Dim varModuleName As Variant
    
    For Each varModuleName In This.ModuleList
        ' TODO Provide a warning if module listed in configuration is not found
        If CheckNameInCollection(varModuleName, This.Project.VBComponents) Then
            Dim comModule As VBComponent
            Set comModule = This.Project.VBComponents(varModuleName)
            
            Dim Dest As String
            Dest = This.Path & Application.PathSeparator & varModuleName & FileExtension(comModule)
            comModule.Export Dest

            If This.FormDeleted Then
                If comModule.Type = vbext_ct_Document Then
                    comModule.CodeModule.DeleteLines 1, comModule.CodeModule.CountOfLines
                Else
                    This.Project.VBComponents.Remove comModule
                End If
            End If
        End If
    Next varModuleName

    '// Remove all references listed
    If This.FormDeleted Then
        For Each varModuleName In This.ModuleList
            If CheckNameInCollection(varModuleName, This.Project.References) Then
                This.Project.References.Remove This.Project.References(varModuleName)
            End If
        Next varModuleName
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
    
    If This.Project.Name = "Personal" Then
        MsgBox "Can not import PERSONAL.xlsb because " & _
               "the import code is in PERSONAL.xlsb", _
               vbOKOnly Or vbCritical, _
               "Can Not Import PERSONAL.xlsb"
        Exit Sub
    End If

    '// Import code from listed module files
    Dim varModuleName As Variant
    Dim comModule As VBComponent
    
    For Each varModuleName In This.ModuleList
        Set comModule = This.Project.VBComponents(varModuleName)
        ImportModule _
            This.Project, _
            varModuleName, _
            This.Path & Application.PathSeparator & _
                This.ModuleList(varModuleName).Module & "." & This.ModuleList(varModuleName).Extension
            
    Next varModuleName

    '// Add references listed in the config file
    Dim Entry As Variant
    Dim Ref As VBAReferences_Table
    
    For Each Entry In This.ReferencesList
        Set Ref = This.ReferencesList(Entry)
        
        If Not CheckNameInCollection(Ref.Name, This.Project.References) Then
            This.Project.References.AddFromGuid _
                GUID:=Ref.GUID, _
                Major:=Ref.Major, _
                Minor:=Ref.Minor
        End If
    Next Entry
    
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
    
    GitForm.Caption = "Select the VBA Project to " & TitleText
    GitForm.Show
    
    Dim SelectedProject As VBProject
    
    If This.FormCanceled Then
        ' Either Cancel button or dialog close button (red X) selected
        MsgBox TitleText & " canceled by user", _
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
    
    Dim collAddList As Collection
    Set collAddList = New Collection
    
    Dim strAddListStr As String
    strAddListStr = vbNullString
    
    Dim comModule As VBComponent
    Dim boolCreateNewEntry  As Boolean
    
    For Each comModule In This.Project.VBComponents
        boolCreateNewEntry = _
                           ExportableModule(comModule) And _
                           Not This.ModuleList.Exists(comModule.Name)

        If boolCreateNewEntry Then
            collAddList.Add comModule.Name
            strAddListStr = strAddListStr & comModule.Name & vbNewLine
        End If
    Next comModule

    ' Ask the user if they want to add new modules to the config file
    Dim intUserResponse As Long
    Dim NewMod As VBAModuleList_Table
    Dim varModuleName As Variant
    
    If collAddList.Count > 0 Then
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some modules not listed in the configuration file which " & _
                          "exist in the current project. Would you like to " & _
                          "add these modules to the configuration file?" & _
                          vbNewLine & _
                          "Note: All modules are listed if there is no existing configuration file" & _
                          vbNewLine & _
                          "New modules:" & vbNewLine & _
                          strAddListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="New Modules")

        If intUserResponse = vbYes Then
            For Each varModuleName In collAddList
                Set NewMod = New VBAModuleList_Table
                If This.ModuleList.Exists(varModuleName) Then
                    ReportWarning "Duplicate module name", "Routine", RoutineName, "Module Name", varModuleName
                Else
                    NewMod.Module = varModuleName
                    NewMod.Extension = FileExtension(This.Project.VBComponents(varModuleName))
                    This.ModuleList.Add varModuleName, NewMod
                End If
            Next varModuleName
        End If
    End If
    
    '// Ask user if they want to delete entries for missing modules
    ' Create the list of modules to potentially delete
    Dim collDeleteList As Collection
    Set collDeleteList = New Collection

    Dim strDeleteListStr As String
    strDeleteListStr = vbNullString

    Dim boolDeleteModule As Boolean
    
    For Each varModuleName In This.ModuleList
        boolDeleteModule = True

        If CheckNameInCollection(varModuleName, This.Project.VBComponents) Then
            If ExportableModule(This.Project.VBComponents(varModuleName)) Then
                boolDeleteModule = False
            End If
        End If

        If boolDeleteModule Then
            collDeleteList.Add varModuleName
            strDeleteListStr = strDeleteListStr & varModuleName & vbNewLine
        End If
    Next varModuleName
    ' Now have a list of modules to potentially delete

    ' Ask the user if they want to delete any modules
    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some modules listed in the configuration file which " & _
                          "haven't been found in the current project. Would you like to " & _
                          "remove these modules from the configuration file?" & _
                          vbNewLine & _
                          vbNewLine & _
                          "Missing modules:" & vbNewLine & _
                          strDeleteListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="Missing Modules")

        If intUserResponse = vbYes Then
            For Each varModuleName In collDeleteList
                This.ModuleList.Remove varModuleName
            Next varModuleName
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
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Choose Configuration File Directory"
        .InitialFileName = InitialDirectory
        Dim Response As Long
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
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Choose Base Path Directory"
        .InitialFileName = InitialDirectory
        Dim Response As Long
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

Private Sub GetReferences(ByVal ThisProject As VBProject)

    ' This routine gathers a list of the references in this project
    ' Compares that list with the existing config file
    ' Modifies the list of reference if the user desires
    
    Const RoutineName As String = Module_Name & "GetReferences"
    On Error GoTo ErrorHandler
    
    Dim collAddList As Collection
    Set collAddList = New Collection
    
    Dim strAddListStr As String
    strAddListStr = vbNullString
    
    Dim refReference As Reference
    For Each refReference In ThisProject.References
        If Not refReference.BuiltIn Then
            If ReferenceToAdd(refReference) Then
                If Not This.ReferencesList.Exists(refReference.Name) Then
                    collAddList.Add refReference
                    strAddListStr = strAddListStr & refReference.Name & vbNewLine
                End If
            End If
        End If
    Next refReference

    ' Ask the user if they want to add new references to the config file
    If collAddList.Count > 0 Then
        Dim intUserResponse As Long
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some references not listed in the configuration file which " & _
                          "exist in the current project. Would you like to " & _
                          "add these references to the configuration file?" & _
                          vbNewLine & _
                          "Note: if the configuration file doesn't already exist, this will be a list of all references" & _
                          vbNewLine & _
                          "New references:" & vbNewLine & _
                          strAddListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="New References")

        If intUserResponse = vbYes Then
            Dim I As Long
            I = 1
            Dim Ref As Variant
            Dim NewRef As VBAReferences_Table
            For Each Ref In collAddList
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
        If Not CheckNameInCollection(Ref, ThisProject.References) Then
            collDeleteList.Add Ref
            strDeleteListStr = strDeleteListStr & Ref & vbNewLine
        End If
    Next Ref

    ' Ask the user if they want to delete any references
    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
                          Prompt:="There are some references listed in the configuration file which " & _
                                   "haven't been found in the current project. Would you like to " & _
                                   "remove these references from the configuration file?" & vbNewLine & _
                                   vbNewLine & _
                                   "Missing references:" & vbNewLine & _
                                   strDeleteListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="Missing References")

        If intUserResponse = vbYes Then
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
        ReferenceToAdd = False                   ' Bad reference; skip
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
    '// Import a VBA code module... how hard could it be right?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ImportModule"
    On Error GoTo ErrorHandler
    
    Dim ErrorNumber As Long
    Dim comNewImport As VBComponent
    
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
    
    Set comNewImport = ThisProject.VBComponents.Import(ModulePath)
    ErrorNumber = Err.Number
    If ErrorNumber = 60061 Then Exit Sub         ' Module already in use
    If ErrorNumber = 53 Then Exit Sub         ' Module does not exist
    On Error GoTo 0
    
    If comNewImport.Name <> ModuleName Then
        If CheckNameInCollection(ModuleName, ThisProject.VBComponents) Then

            Dim comExistingComp As VBComponent
            Set comExistingComp = ThisProject.VBComponents(ModuleName)
            If comExistingComp.Type = vbext_ct_Document Then

                Dim modCodeCopy As CodeModule
                Set modCodeCopy = comNewImport.CodeModule
                
                Dim modCodePaste As CodeModule
                Set modCodePaste = comExistingComp.CodeModule
                
                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                
                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.lines(1, modCodeCopy.CountOfLines)
                End If
                
                ThisProject.VBComponents.Remove comNewImport

            Else
                comExistingComp.Name = comExistingComp.Name & "_remove"
                ThisProject.VBComponents.Remove comExistingComp
                comNewImport.Name = ModuleName   ' TODO fails on work computer
                ThisProject.VBComponents.Remove comExistingComp
            End If
        Else

            comNewImport.Name = ModuleName

        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ImportModule

Private Function ExportableModule(ByVal comModule As VBComponent) As Boolean
    '// Is the given module exportable by this tool?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ExportableModule"
    On Error GoTo ErrorHandler
    
    ExportableModule = _
                     (Not ModuleEmpty(comModule)) _
                     And _
                     (FileExtension(comModule) <> vbNullString)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ExportableModule

Private Function ModuleEmpty(ByVal comModule As VBComponent) As Boolean
    '// Check if a code module is effectively empty.
    '// effectively empty should be functionally and semantically equivalent to
    '// actually empty.
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ModuleEmpty"
    On Error GoTo ErrorHandler
    
    ModuleEmpty = True

    Dim lngNumLines As Long
    lngNumLines = comModule.CodeModule.CountOfLines
    
    Dim lngCurLine As Long
    For lngCurLine = 1 To lngNumLines
        Dim strCurLine As String
        strCurLine = comModule.CodeModule.lines(lngCurLine, 1)
        
        If Not (strCurLine = "Option Explicit" Or strCurLine = vbNullString) Then
            ModuleEmpty = False
            Exit Function
        End If
    Next lngCurLine
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ModuleEmpty

Private Function FileExtension(ByVal comModule As VBComponent) As String
    '// The appropriate file extension for exporting the given module
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "FileExtension"
    On Error GoTo ErrorHandler
    
    Select Case comModule.Type
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
    ' Version 1.0
    ' Added error handling
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "EnsurePath"
    On Error GoTo ErrorHandler
    
    Set FSO = New Scripting.FileSystemObject
    
    Dim strParentPath As String
    strParentPath = FSO.GetParentFolderName(Path)

    If strParentPath <> vbNullString Then
        EnsurePath strParentPath
        If Not FSO.FolderExists(strParentPath) Then
            If FSO.FileExists(strParentPath) Then
                ReportError "No path exists", _
                            "Path", strParentPath
            Else
                FSO.CreateFolder (strParentPath)
            End If
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' EnsurePath

Public Sub LetGitFormCanceled(ByVal Canx As Boolean)
    This.FormCanceled = Canx
End Sub                                          ' LetGitFormCanceled

Public Sub LetGitFormDelete(ByVal DeleteCheck As Boolean)
    If DeleteCheck Then
        Select Case MsgBox( _
               "Are you sure you want to delete all the code in this project when you export?", _
               vbYesNo Or vbExclamation, _
               "Delete All Code?")
        Case vbYes
            This.FormDeleted = True
        Case vbNo
            This.FormDeleted = False
        End Select
    End If
End Sub                                          ' LetGitFormDelete

Private Sub PopulateTables(ByVal Title As String)

    ' Reads all the data from the tables
    
    Const RoutineName As String = Module_Name & "PopulateTables"
    On Error GoTo ErrorHandler
    
    GetProject Title, This.Project, This.Workbook
    
    Set This.Worksheet = This.Workbook.Worksheets(VBAMakeFile)
    
    Dim ModuleList As VBAModuleList_Table
    Set ModuleList = New VBAModuleList_Table
    
    Set This.ModuleTable = This.Worksheet.ListObjects(VBAModuleList)
    
    If Table.TryCopyTableToDictionary( _
        ModuleList, This.ModuleList, This.ModuleTable, False) _
    Then
    Else
        ReportError "Error loading Module List", "Routine", RoutineName
        GoTo Done
    End If
    
    Dim SourceFolder As VBASourceFolder_Table
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
    
    Dim RefList As VBAReferences_Table
    Set RefList = New VBAReferences_Table
    
    Set This.ReferencesTable = This.Worksheet.ListObjects(VBAReferences)
    
    If Table.TryCopyTableToDictionary( _
        RefList, This.ReferencesList, This.ReferencesTable, False) _
    Then
    Else
        ReportError "Error loading References List", "Routine", RoutineName
        GoTo Done
    End If
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' PopulateTables

