VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaveComponents 
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2400
   OleObjectBlob   =   "frmSaveComponents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaveComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    frmSaveComponents.chkbxModules.Value = True
    frmSaveComponents.chkbxClasses.Value = True
    frmSaveComponents.chkbxForms.Value = False
End Sub

Private Sub cmdButton_Click()
    SaveModules = frmSaveComponents.chkbxModules.Value
    SaveClasses = frmSaveComponents.chkbxClasses.Value
    SaveForms = frmSaveComponents.chkbxForms.Value
    
    Unload Me
End Sub
