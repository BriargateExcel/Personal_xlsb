Attribute VB_Name = "CodeCommon"
Option Explicit
' Changes
'   Requires references to:
'       Microsoft Visual Basic for Applications Extensibility 5.3
'       Microsoft Scripting Runtime
'
Private Const Module_Name As String = "CodeCommon."

' Concept
'
' VBA Make File Worksheet
'   VBAModule table
'       Module column: all the modules to be manipulated (may need manual augmentation)
'       Extension column: bas, cls, or frm depending on the module
'       Paths column: comma separated list mapping the module to where to store the module
'           Needs manual manipulation
'           All: all the modules in this workbook
'           Common: all the common modules shared across workbooks
'           Built: the modules built the TableBuilder for this workbook
'   VBASourceFolder table
'       Path: a folder path name
'       Path Name: maps the Path in the VBASourceFolder to the Module in the VBAModuleList table
'   VBAReferences table
'       Name: name of the reference
'       Description:
'       GUID:
'       Major:
'       Minor:
'
' GetModules: routine in CodeCommon module
'   Creates the list of VBAProjects
'   Populates and displays the list of VBAProjects
'   Lets the user pick the VBAProject
'   Creates the list of ModuleGroups based on the selected VBAProject
'   Populates and displays the list of module groups
'   Lets the user pick the ModuleGroup
'   Returns the VBAProject, ModuleGroup, and Path
'
' CodeMakeConfigTables module
'   Build the configuration tables for this workbook
'   Add the worksheet if missing
'   Add the tables if missing
'   Populate the talbes to the extent possible automatically
'   Reminds the user to manually populate the remaining data
'
' CodeExport module
'   Calls GetModules to get VBAProject, ModuleGroup, and Path
'   Exports the modules to the proper path
'
' CodeImport module
'   Calls GetModules to get VBAProject, ModuleGroup, and Path
'   Imports the modules from the proper path
'

Private Type CodeType
    Title As String
    
    FormCanceled As Boolean
    
    Workbook As Workbook
    ProjectName As String
    Project As VBProject
    
    ModuleDict As Dictionary
    ModuleTable As ListObject
    
    ModuleGroup As String
    ModuleGroupArray As Variant
    ModuleGroupDict As Dictionary
    
    FolderPath As String
    FolderPathDict As Dictionary
    FolderPathTable As ListObject
    
    ReferenceTable As ListObject
    ReferenceDict As Dictionary
End Type

Private This As CodeType

Private Const VBAReferences As String = "VBAReferences"

Public Property Get ModuleTable() As ListObject
    Set ModuleTable = This.ModuleTable
End Property

Public Property Get ModuleDict() As Dictionary
    Set ModuleDict = This.ModuleDict
End Property

Public Property Get ReferenceTable() As ListObject
    Set ReferenceTable = This.ReferenceTable
End Property

Public Property Get ReferenceDict() As Dictionary
    Set ReferenceDict = This.ReferenceDict
End Property

Public Property Set ReferenceDict(ByVal Param As Dictionary)
    Set This.ReferenceDict = Param
End Property

Public Property Get ModuleGroupDict() As Dictionary
    Set ModuleGroupDict = This.ModuleGroupDict
End Property

Public Property Get FolderPathTable() As ListObject
    Set FolderPathTable = This.FolderPathTable
End Property

Public Property Get FolderPathDict() As Dictionary
    Set FolderPathDict = This.FolderPathDict
End Property

Public Property Get FolderPath() As String
    FolderPath = This.FolderPath
End Property

Public Property Let FolderPath(ByVal Vbl As String)
    This.FolderPath = Vbl
End Property

Public Property Get ModuleGroupArray() As Variant
    ModuleGroupArray = This.ModuleGroupArray
End Property

Public Property Get ModuleGroup() As String
    ModuleGroup = This.ModuleGroup
End Property

Public Sub SetTitle(ByVal Title As String)
    This.Title = Title
End Sub

Public Sub SetProjectName(ByVal ProjectName As String)
    This.ProjectName = ProjectName
End Sub

Public Sub SetModuleGroup(ByVal ModuleGroup As String)
    This.ModuleGroup = ModuleGroup
End Sub

Public Property Get WorkBookName() As String
    WorkBookName = This.Workbook.Name
End Property

Public Property Get VBAProject() As VBProject
    Set VBAProject = This.Project
End Property

Public Sub LetGitFormCanceled(ByVal Canx As Boolean)
    This.FormCanceled = Canx
End Sub

Public Property Get GetFormCanceled()
    GetFormCanceled = This.FormCanceled
End Property

Public Sub GetModulesOfInterest()

    ' Builds an array of the modules of interest to the user
    
    Const RoutineName As String = Module_Name & "GetModulesOfInterest"
    On Error GoTo ErrorHandler
    
    InitializeStep1
        
    If GetFormCanceled Then GoTo Done
    
    GetFolderPath
    
    ' Using the selected group and the list of modules in This.ModuleDict
    ' Select the modules of interest
    Dim VBAM As Dictionary
    Set VBAM = New Dictionary
    
    Dim Module As VBAModuleList_Table
    Dim GroupArray As Variant
    Dim Element As String
    Dim NumModules As Long
    Dim Entry As Variant
    Dim I As Long
    
    For Each Entry In This.ModuleDict
        Set Module = This.ModuleDict(Entry)
        GroupArray = Split(Module.Paths, ",")
        
        For I = LBound(GroupArray, 1) To UBound(GroupArray, 1)
            Element = Trim$(GroupArray(I))
            If Element = This.ModuleGroup Then
                NumModules = NumModules + 1
            End If
        Next I
    Next Entry
    ' Now know how many modules of interest
    
    ReDim This.ModuleGroupArray(1 To NumModules)
    
    Dim J As Long
    J = 1
    
    Set This.ModuleGroupDict = New Dictionary
    
    For Each Entry In This.ModuleDict
        Set Module = This.ModuleDict(Entry)
        GroupArray = Split(Module.Paths, ",")
        
        For I = LBound(GroupArray, 1) To UBound(GroupArray, 1)
            Element = Trim(GroupArray(I))
            If Element = This.ModuleGroup Then
                This.ModuleGroupDict.Add Module.Module, Module
                This.ModuleGroupArray(J) = Module.Module
                J = J + 1
            End If
        Next I
    Next Entry
    ' Now have an array of the modules of interest
    
Done:
    Exit Sub
ErrorHandler:
    MsgBox "Exception raised" & vbCrLf & _
                "Routine: " & RoutineName & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Description: " & Err.Description
End Sub ' GetModulesOfInterest

Public Function FileExtension(ByVal ThisModule As VBComponent) As String
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
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' FileExtension

Private Function WorkbookOfVBProject(WhichVBP As Variant) As Workbook

    ' Used as the top level routine
    
    Const RoutineName As String = Module_Name & "WorkbookOfVBProject"
    On Error GoTo ErrorHandler
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' From Chip Pearson
' WorkbookOfVBProject
' This returns the Workbook object for a specified VBIDE.VBProject.
' The parameter WhichVBP can be any of the following:
'   A VBIDE.VBProject object
'   A string containing the name of the VBProject.
'   The index number (ordinal position in Project window) of the VBProject.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WB As Workbook
Dim AI As AddIn
Dim VBP As VBIDE.VBProject

If IsObject(WhichVBP) = True Then
    ' If WhichVBP is an object, it must be of the
    ' type VBIDE.VBProject. Any other object type
    ' throws an error 13 (type mismatch).
    On Error GoTo 0
    If TypeOf WhichVBP Is VBIDE.VBProject Then
        Set VBP = WhichVBP
    Else
        Err.Raise 13
    End If
Else
    On Error Resume Next
    Err.Clear
    ' Here, WhichVBP is either the string name of
    ' the VBP or its ordinal index number.
    Set VBP = Application.VBE.VBProjects(WhichVBP)
    On Error GoTo 0
    If VBP Is Nothing Then
        Err.Raise 9
    End If
End If

For Each WB In Workbooks
    If WB.VBProject Is VBP Then
        Set WorkbookOfVBProject = WB
        GoTo Done
    End If
Next WB
' not found in workbooks, search installed add-ins.
For Each AI In Application.AddIns
    If AI.Installed = True Then
        If Workbooks(AI.Name).VBProject Is VBP Then
            Set WorkbookOfVBProject = Workbooks(AI.Name)
            GoTo Done
        End If
    End If
Next AI
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' WorkbookOfVBProject

Private Sub GetFolderPath()

    ' Used as the top level routine
    
    Const RoutineName As String = Module_Name & "GetFolderPath"
    On Error GoTo ErrorHandler
        
    Dim VBAProj As VBIDE.VBProject
    Set VBAProj = Application.VBE.VBProjects(This.ProjectName)
    
    Set This.FolderPathTable = Workbooks(WorkBookName).Worksheets("VBA Make File").ListObjects("VBASourceFolder")
    
    Dim VBAMType As VBASourceFolder_Table
    Set VBAMType = New VBASourceFolder_Table
    
    Set This.FolderPathDict = New Dictionary
    
    If TryCopyTableToDictionary(VBAMType, This.FolderPathDict, This.FolderPathTable, False) Then
    Else
        ReportError "Error copying paths table to dictionary", "Routine", RoutineName
        GoTo Done
    End If
    ' Now have a dictionary of the VBASourceFolder table we want
    
    Dim Entry As Variant
    Dim PathEntry As VBASourceFolder_Table
    
    For Each Entry In This.FolderPathDict
        Set PathEntry = This.FolderPathDict(Entry)
        
        If PathEntry.PathName = This.ModuleGroup Then
            This.FolderPath = PathEntry.Path
        End If
    Next Entry
    
    EnsurePath This.FolderPath
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' GetFolderPath

Private Sub InitializeStep1()

    ' Fetch the name of the project the user is interested in
    
    Const RoutineName As String = Module_Name & "InitializeStep1"
    On Error GoTo ErrorHandler
    
    ' Clear out all the dictionaries
    Set This.FolderPathDict = Nothing
    Set This.ModuleDict = Nothing
    Set This.ModuleGroupDict = Nothing
    Set This.ReferenceDict = Nothing
    
    ' Display a list of the VBAProjects currently active to the user
    Dim ProjList As Variant
    ReDim ProjList(1 To Application.VBE.VBProjects.Count)
    
    Dim Entry As Variant
    Dim I As Long
    I = 1
    For Each Entry In Application.VBE.VBProjects
        ProjList(I) = Entry.Name
        I = I + 1
    Next Entry
    GitForm.AddProjectList This.Title, ProjList
    ' The user has now selected a project name
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
    CloseErrorFile
End Sub ' InitializeStep1
    
Public Sub InitializeStep2()
' 7/5/2020
'   Restructured fetching of VBAModulesList

    ' Gather all the data from the three tables
    ' Check the data for consistency
    
    Const RoutineName As String = Module_Name & "InitializeStep2"
    On Error GoTo ErrorHandler
    
    Set This.Project = Application.VBE.VBProjects(This.ProjectName)
    Set This.Workbook = WorkbookOfVBProject(This.Project)
    
    ' Get the modules table
    Dim ErrorNumber As Long
    On Error Resume Next
    Set This.ModuleTable = Workbooks(WorkBookName).Worksheets("VBA Make File").ListObjects("VBAModuleList")
    On Error GoTo ErrorHandler
    ErrorNumber = Err.Number
    If ErrorNumber <> 0 Then
        ReportError "VBAModuleList not found", "Routine", RoutineName
        GoTo ErrorHandler
    End If
    
    Dim Module As VBAModuleList_Table
    Set Module = New VBAModuleList_Table
    
    If Table.TryCopyTableToDictionary(Module, This.ModuleDict, This.ModuleTable, False) Then
    Else
        ReportError "Error Copying Modules table", "Routine", RoutineName
        GoTo ErrorHandler
    End If
    ' Now we have the modules table
    
    Dim GroupDict As Dictionary
    Set GroupDict = New Dictionary
    
    Dim Entry As Variant
    Dim PathList As String
    Dim PathArray As Variant
    Dim I As Long
    Dim Element As String
    For Each Entry In This.ModuleDict
        PathList = This.ModuleDict(Entry).Paths
        PathArray = Split(PathList, ",")
        
        For I = LBound(PathArray, 1) To UBound(PathArray, 1)
            Element = Trim(PathArray(I))
            If Not GroupDict.Exists(Element) Then
                GroupDict.Add Element, Element
            End If
        Next I
    Next Entry
    ' Now we have a dictionary of the group names
    
    Dim GroupArray As Variant
    ReDim GroupArray(1 To GroupDict.Count)
    
    I = 1
    For Each Entry In GroupDict
        GroupArray(I) = Entry
        I = I + 1
    Next Entry
    ' Now we have an array of the group names
    
    ' Get the paths table
    On Error Resume Next
    Set This.FolderPathTable = Workbooks(WorkBookName).Worksheets("VBA Make File").ListObjects("VBASourceFolder")
    On Error GoTo ErrorHandler
    ErrorNumber = Err.Number
    If ErrorNumber <> 0 Or This.FolderPathTable Is Nothing Then
        MsgBox "Paths table does not exist"
        GoTo Done
    Else
        Dim Paths As VBASourceFolder_Table
        Set Paths = New VBASourceFolder_Table
        
        If Table.TryCopyTableToDictionary(Paths, This.FolderPathDict, This.FolderPathTable, False) Then
        Else
            ReportError "Error Copying Paths table", "Routine", RoutineName
            GoTo ErrorHandler
        End If
    End If
    ' Now we have the paths table
    
    For Each Entry In GroupDict
        If Not This.FolderPathDict.Exists(Entry) Then
            GitForm.CancelProcessing
            ReportError "Missing an entry in the Paths table for " & Entry, "Routine", RoutineName
            LetGitFormCanceled True
            GoTo ErrorHandler
        End If
    Next Entry
    
    ' Get the references table
    On Error Resume Next
    Set This.ReferenceTable = Workbooks(WorkBookName).Worksheets("VBA Make File").ListObjects("VBAReferences")
    On Error GoTo ErrorHandler
    ErrorNumber = Err.Number
    If ErrorNumber <> 0 Or This.ReferenceTable Is Nothing Then
        MsgBox "References table does not exist"
        GoTo Done
    Else
        If This.ReferenceTable.ListRows.Count = 0 Then
        Else
            Dim Ref As VBAReferences_Table
            Set Ref = New VBAReferences_Table
            
            If Table.TryCopyTableToDictionary(Ref, This.ReferenceDict, This.ReferenceTable, False) Then
            Else
                ReportError "Error Copying References table", "Routine", RoutineName
                GoTo Done
            End If
        End If
    End If
    ' Now we have all the tables
    
    GitForm.AddModuleGroupList GroupArray
    
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
    CloseErrorFile
End Sub ' InitializeStep2

