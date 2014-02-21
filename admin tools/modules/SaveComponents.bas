Attribute VB_Name = "SaveComponents"
Option Explicit
Private objMyProj As VBProject
Private objVBComp As VBComponent
Private FileExtension$
Private SubDirectory$
Private DirectoryError As Boolean

Public SaveModules As Boolean
Public SaveClasses As Boolean
Public SaveForms As Boolean

Sub EntryPoint()
    Call InitializeObjects
    Call OpenForm
    Call EvaluateComponents
    If DirectoryError Then
        MsgBox ("Check directory exists!")
    End If
End Sub

Private Sub InitializeObjects()
    Set objMyProj = Application.VBE.VBProjects("MacroTools") 'SET TO PROJECT
    DirectoryError = False
End Sub

Private Sub OpenForm()
    frmSaveComponents.Show
End Sub

Private Sub EvaluateComponents()
    For Each objVBComp In objMyProj.VBComponents
        If IsClassModuleOrForm(objVBComp) Then
            Call SaveComponent(objVBComp)
        End If
    Next
End Sub
Private Function IsClassModuleOrForm(component As VBComponent) As Boolean
    IsClassModuleOrForm = False
    
    Select Case component.Type
        Case vbext_ct_StdModule
            If SaveModules Then
                Call SetExtensionAndSubDirectory("modules\", ".bas")
                IsClassModuleOrForm = True
            End If
        Case vbext_ct_ClassModule
            If SaveClasses Then
                Call SetExtensionAndSubDirectory("classes\", ".cls")
                IsClassModuleOrForm = True
            End If
        Case vbext_ct_MSForm
            If SaveForms Then
                Call SetExtensionAndSubDirectory("forms\", ".frm")
                IsClassModuleOrForm = True
            End If
    End Select
End Function
Private Sub SetExtensionAndSubDirectory(sub_directory$, file_extension$)
    FileExtension = file_extension
    SubDirectory = sub_directory
End Sub
Private Sub SaveComponent(component As VBComponent)
    Dim directory$, filename$
    
    directory = "C:\Users\dl_2\Documents\Backups\GFC Tools\admin tools\" & SubDirectory 'SET TO PRIMARY DIRECTORY
    filename = component.Name & FileExtension
    
    On Error GoTo Message
    component.Export directory & filename
    Exit Sub
    
Message:
    DirectoryError = True
End Sub
