Attribute VB_Name = "CodeImportExportMenus"
Option Explicit

Private Const Module_Name As String = "CodeIMportExportMenus."

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection
Private Const CommandBarName As String = "Personal Routines"

Private Type CodeType
    CustomBar As CommandBar
End Type

Private This As CodeType

Public Sub Auto_Open()
' https://bettersolutions.com/vba/ribbon/face-ids-2003.htm for FaceIDs
    
    Dim NewButton As CommandBarButton
    
    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0
    
    Set This.CustomBar = CommandBars.Add(Name:=CommandBarName)
    
    BuildButton "MakeConfigurationTables", "Make Configuration Tables", 538
    BuildButton "Export", "Export", 7026
    BuildButton "Import", "Import", 7027
    BuildButton "ExposeAllSheets", "Expose All Sheets", 703
    
    This.CustomBar.Visible = True
    
End Sub ' Auto_Open

Private Sub BuildButton( _
    ByVal RoutineToExecute As String, _
    ByVal Caption As String, _
    ByVal FaceID As Long)

    ' Build one button on the command bar
    
    Const RoutineName As String = Module_Name & "BuildButton"
    On Error GoTo ErrorHandler
    
    Dim NewButton As CommandBarButton
    
    Set NewButton = This.CustomBar.Controls.Add(msoControlButton)
    NewButton.OnAction = RoutineToExecute
    NewButton.Caption = Caption
    NewButton.FaceID = FaceID
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' BuildButton

Public Sub Auto_Close()
    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0
End Sub ' Auto_Close
