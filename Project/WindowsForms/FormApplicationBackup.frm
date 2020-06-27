VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormApplicationBackup 
   Caption         =   "Assistente de BackUp"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "FormApplicationBackup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormApplicationBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InitialFileName As String


Private Sub ButtonSelectBackupLocation_Click()
    TextDirectoryPath.Text = SysFunction.SelectFolder()
End Sub

Private Sub UserForm_Initialize()
    
    On Error GoTo Exception
      
      Call FormDefination
    
    Exit Sub
    
Exception:
    
    Call SysMethod.SubmitException
    
End Sub


Private Sub FormDefination()

    Call SysMethod.DefineUserFormStyle(Me)
    
    InitialFileName = SysDirectory.PathSheet & "\User\Backup"
    
    TextDirectoryPath.Text = InitialFileName

End Sub
