Attribute VB_Name = "VBASourceFolder"
Option Explicit

' Built on 6/15/2020 10:41:18 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "VBASourceFolder."

Private Type PrivateType
    Initialized As Boolean
    Dict As Dictionary
    Wkbk As Workbook
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Private Const pPathColumn As Long = 1
Private Const pPathNameColumn As Long = 2
Private Const pHeaderWidth As Long = 2

Private Const pFileName As String = "Blank"
Private Const pWorksheetName As String = vbNullString
Private Const pExternalTableName As String = vbNullString

Public Property Get PathColumn() As Long
    PathColumn = pPathColumn
End Property ' PathColumn

Public Property Get PathNameColumn() As Long
    PathNameColumn = pPathNameColumn
End Property ' PathNameColumn

Public Property Get Headers() As Variant
    Headers = Array( _
        "Path", _
        "Path Name")
End Property ' Headers

Public Property Get Dict() As Dictionary
   Set Dict = This.Dict
End Property ' Dict

Public Property Get SpecificTable() As ListObject
    ' This table is handled in other ways
    Set SpecificTable = Nothing
End Property ' SpecificTable

Public Property Get Initialized() As Boolean
   Initialized = This.Initialized
End Property ' Initialized

Public Sub Initialize()

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    Dim LocalTable As VBASourceFolder_Table
    Set LocalTable = New VBASourceFolder_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(LocalTable, This.Dict, VBASourceFolder.SpecificTable) Then
        This.Initialized = True
    Else
        ReportError "Error copying VBASourceFolder table", "Routine", RoutineName
        This.Initialized = False
        GoTo Done
    End If

    If Not This.Wkbk Is Nothing Then This.Wkbk.Close
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' VBASourceFolderInitialize

Public Sub Reset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub ' Reset

Public Property Get HeaderWidth() As Long
    HeaderWidth = pHeaderWidth
End Property ' HeaderWidth

Public Property Get GetPathFromPathName(ByVal PathName As String) As String

    Const RoutineName As String = Module_Name & "GetPathFromPathName"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then VBASourceFolder.Initialize

    If CheckPathNameExists(PathName) Then
        GetPathFromPathName = This.Dict(PathName).Path
    Else
        ReportError "Unrecognized PathName", _
            "Routine", RoutineName, _
            "Path Name", PathName
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Path Name", PathName

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetPathFromPathName

Public Function CreateKey(ByVal Record As VBASourceFolder_Table) As String

    Const RoutineName As String = Module_Name & "CreateKey"
    On Error GoTo ErrorHandler

    CreateKey = Record.PathName

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CreateKey

Public Function TryCopyDictionaryToArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyDictionaryToArray"
    On Error GoTo ErrorHandler

    TryCopyDictionaryToArray = True

    If Dict.Count = 0 Then
        ReportError "Error copying VBASourceFolder_Table dictionary to array,", "Routine", RoutineName
        TryCopyDictionaryToArray = False
        GoTo Done
    End If

    ReDim Ary(1 To Dict.Count, 1 To 2)

    Dim I As Long
    I = 1

    Dim Record As VBASourceFolder_Table
    Dim Entry As Variant
    For Each Entry In Dict.Keys
        Set Record = Dict.Item(Entry)

        Ary(I, pPathColumn) = Record.Path
        Ary(I, pPathNameColumn) = Record.PathName

        I = I + 1
    Next Entry

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' VBASourceFolderTryCopyDictionaryToArray

Public Function TryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As VBASourceFolder_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Set Record = New VBASourceFolder_Table

            Record.Path = Ary(I, pPathColumn)
            Record.PathName = Ary(I, pPathNameColumn)

            Key = VBASourceFolder.CreateKey(Record)

            If Not Dict.Exists(Key) Then
                Dict.Add Key, Record
            Else
                ReportWarning "Duplicate key", "Routine", RoutineName, "Key", Key
                TryCopyArrayToDictionary = False
                GoTo Done
            End If
        Next I

    Else
        ReportError "Invalid Array", "Routine", RoutineName
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' VBASourceFolderTryCopyArrayToDictionary

Public Function CheckPathNameExists(ByVal PathName As String) As Boolean _

    Const RoutineName As String = Module_Name & "CheckPathNameExists"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then VBASourceFolder.Initialize

    If PathName = vbNullString Then
        CheckPathNameExists = True
        GoTo Done
    End If

    CheckPathNameExists = This.Dict.Exists(PathName)

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Path Name", PathName

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckPathNameExists

Public Sub FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "VBASourceFolderFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler


Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' VBASourceFolderFormatArrayAndWorksheet

' No application unique routines found

