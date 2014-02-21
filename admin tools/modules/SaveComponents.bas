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
    Set objMyProj = Application.VBE.VBProjects("MacroTools")
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
    IsClassModuleOrForm = True
    
    Select Case component.Type
        Case vbext_ct_StdModule
            Call SetExtensionAndSubDirectory("modules\", ".bas")
        Case vbext_ct_ClassModule
            Call SetExtensionAndSubDirectory("classes\", ".cls")
        Case vbext_ct_MSForm
            Call SetExtensionAndSubDirectory("forms\", ".frm")
        Case Else
            IsClassModuleOrForm = False
    End Select
End Function
Private Sub SetExtensionAndSubDirectory(sub_directory$, file_extension$)
    FileExtension = file_extension
    SubDirectory = sub_directory
End Sub
Private Sub SaveComponent(component As VBComponent)
    Dim directory$, filename$
    
    directory = "C:\Users\dl_2\Documents\Backups\GFC Tools\admin tools\" & SubDirectory
    filename = component.Name & FileExtension
    
    On Error GoTo Message
    component.Export directory & filename
    Exit Sub
    
Message:
    DirectoryError = True
End Sub
