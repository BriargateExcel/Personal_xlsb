VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitForm 
   Caption         =   "Select the VBA Project to Export/Import"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   OleObjectBlob   =   "GitForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSelectingProject As Boolean
Private pModuleGroup As Variant

Private Sub CancelButton_Click()
    LetGitFormCanceled True
    Me.Hide
End Sub

Private Sub ModuleGroupSelectButton_Click()
    LetGitFormCanceled False
    SetModuleGroup Me.ModuleGroupList.Value
    Me.Hide
End Sub

Private Sub ProjectSelectButton_Click()
    LetGitFormCanceled False
    
    SetProjectName Me.ProjectList.Value
    
    Me.ModuleGroupList.Visible = True
    Me.ModuleGroupSelectButton.Visible = True
    
    InitializeStep2
End Sub

Private Sub UserForm_Activate()
    
    LetGitFormCanceled False
    
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

    Me.ModuleGroupList.Visible = False
    Me.ModuleGroupSelectButton.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    LetGitFormCanceled True
End Sub

Private Sub UserForm_Terminate()
    LetGitFormCanceled True
End Sub

Public Sub AddProjectList( _
    ByVal Title As String, _
    ByVal ProjList As Variant)
    
    pSelectingProject = True
    
    Me.Caption = Title
    
    Me.ProjectList.Clear
    Me.ModuleGroupList.Clear
    
    Dim I As Long
    For I = 1 To UBound(ProjList, 1)
        Me.ProjectList.AddItem ProjList(I)
    Next I
    
    If ProjList(1) = "Personal" Then
        Me.ProjectList.Value = ProjList(2)
    Else
        Me.ProjectList.Value = ProjList(1)
    End If
    
    Me.Show
    
End Sub

Public Sub AddModuleGroupList(ByVal ModGrpList As Variant)
    Me.ModuleGroupList.Clear
    
    Dim I As Long
    For I = 1 To UBound(ModGrpList, 1)
        Me.ModuleGroupList.AddItem ModGrpList(I)
    Next I
    
    Me.ModuleGroupList.Value = ModGrpList(1)
    
End Sub

Public Sub CancelProcessing()
    LetGitFormCanceled True
    Me.Hide
End Sub
