Attribute VB_Name = "CodeImportExportMenus"
'
'
'
' From: https://github.com/spences10/VBA-IDE-Code-Export
'
'
'
Option Explicit

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection

Private Sub CreateXLMenu()
' https://bettersolutions.com/vba/ribbon/face-ids-2003.htm for FaceIDs
    Dim NewButton As CommandBarButton
    
    On Error Resume Next
    CommandBars("Code Manager").Delete
    On Error GoTo 0
    
    Dim CustomBar As CommandBar
    Set CustomBar = CommandBars.Add(Name:="Code Manager")
    
    Set NewButton = CustomBar.Controls.Add(msoControlButton)
    NewButton.OnAction = "MakeConfigFile"
    NewButton.Caption = "Make Config File"
    NewButton.FaceId = 538
    NewButton.TooltipText = "Create or overwrite an existing json file directing which components to export or import"
        
    Set NewButton = CustomBar.Controls.Add(msoControlButton)
    NewButton.OnAction = "Export"
    NewButton.Caption = "Export"
    NewButton.FaceId = 360
    NewButton.TooltipText = "Export the components based on the json file"
        
    Set NewButton = CustomBar.Controls.Add(msoControlButton)
    NewButton.OnAction = "Import"
    NewButton.Caption = "Import"
    NewButton.FaceId = 359
    NewButton.TooltipText = "Import the components in a json file overwriting all existing components of the same name"
        
    CustomBar.Visible = True
    
End Sub ' CreateXLMenu

Private Sub RemoveVBEAndXLMenus()

    On Error Resume Next

    Application.VBE.CommandBars(1).Controls("Custom").Delete

    '// Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing

    Application.CommandBars("Worksheet Menu Bar").Controls("Code Manager").Delete
    On Error GoTo 0

End Sub                                          ' RemoveVBEAndXLMenus

Public Sub auto_open()
    '    CreateVBEMenu
    CreateXLMenu
End Sub                                          ' auto_open

Public Sub auto_close()
    RemoveVBEAndXLMenus
End Sub                                          ' auto_close

Private Sub MenuEvents(ByVal objMenuItem As Object)

    Set MnuEvt = New clsVBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt

End Sub                                          ' MenuEvents

Public Sub btnMakeConfig_onAction(control As IRibbonControl)
    MakeConfigFile
End Sub                                          ' btnMakeConfig_onAction

Public Sub btnExport_onAction(control As IRibbonControl)
    Export
End Sub                                          ' btnExport_onAction

Public Sub btnImport_onAction(control As IRibbonControl)
    Import
End Sub                                          ' btnImport_onAction

