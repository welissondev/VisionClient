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

Private SelectedFolderPath As String

Private Sub UserForm_Initialize()
    
    On Error GoTo Exception
      
        Call SysMethod.DefineUserFormStyle(Me)
        
        SelectedFolderPath = GetSelectedPath()
        TextSelectedFolderPath = SelectedFolderPath
        
    Exit Sub
    
Exception:
    
    Call SysMethod.SubmitException
    
End Sub


Private Sub ButtonSelectBackupLocation_Click()
    
    On Error GoTo Exception
    
        SelectedFolderPath = SysFunction.SelectFolder()
        
        If SelectedFolderPath <> Empty Then
            TextSelectedFolderPath.Text = SelectedFolderPath
        End If
    
    Exit Sub
    
Exception:
    
    Call SysMethod.SubmitException
    
End Sub


Private Sub ButtonGenerateBackup_Click()
    
    On Error GoTo Exception
    
        Dim BackupFileName As String
        Dim PathToCopyFile As String
        
        If TextSelectedFolderPath.Text = Empty Then
              MsgBox "Selecione um diretório para salvar o backup!", _
              vbExclamation, "SELECIONE UMA PASTA"
            Exit Sub
        End If
         
        BackupFileName = "Backup-" & Format(Now(), "dd.MM.yyyy_hh.mm.ss") & "_" & SysProperty.AppName
        PathToCopyFile = SysDirectory.PathSheet
        
        MousePointer = fmMousePointerAppStarting
        
        With New FileZipper
            
            .FileName = BackupFileName
            .SourcePath = PathToCopyFile
            .DestinationPath = SelectedFolderPath
            
            If .ZipFile() = True Then
                
                Call SaveSelectedPath
                
                MsgBox "Backup gerado com sucesso!", vbInformation, "SUCESSO"
                      
            End If
            
        End With
        
      MousePointer = fmMousePointerDefault
    
    Exit Sub
    
Exception:
    
    MousePointer = fmMousePointerDefault
    Call SysMethod.SubmitException
    
End Sub

Private Sub SaveSelectedPath()

    Dim FileName As String
    Dim TextFile As Object

    With New FileSystemObject
    
        FileName = SysDirectory.PathUserDef & "\BackupPath.txt"
                 
        Set TextFile = .OpenTextFile(FileName, 2, True)
          With TextFile
              .WriteLine SelectedFolderPath
              .Close
          End With
              
    End With
    
End Sub
Private Function GetSelectedPath() As String

    Dim FileName As String
    Dim TextFile As Object
    
    With New FileSystemObject
    
        FileName = SysDirectory.PathUserDef & "\BackupPath.txt"
            
        Set TextFile = .OpenTextFile(FileName, 1, True)
              While Not TextFile.AtEndOfStream
                  GetSelectedPath = TextFile.ReadLine
              Wend
          TextFile.Close
             
    End With
    
End Function




