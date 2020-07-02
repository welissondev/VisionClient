Attribute VB_Name = "SysFunctions"
Option Explicit

Public Function SelectFolder(Optional ByVal Title As String = Empty, _
Optional ByVal InitialFileName As String = Empty, Optional ByVal _
AllowMultiSelect As Boolean = False) As String
    
      With Application.FileDialog(msoFileDialogFolderPicker)
          
          .Title = Title
          .InitialFileName = InitialFileName
          .AllowMultiSelect = AllowMultiSelect
          .Show
          
          If .SelectedItems.Count > 0 Then
              SelectFolder = .SelectedItems(1)
          End If
             
      End With
    
End Function

Public Function CreateFolder(ByVal Path As String) As Folder
   
   With New FileSystemObject
   
      If .FolderExists(Path) = False Then
         Set CreateFolder = .CreateFolder(Path)
      End If
      
   End With
   
End Function

Public Function CheckFolderExists(FileSpec As String, Optional ByVal Create As Boolean = False) As Boolean
   
   Dim ExistsFile As Boolean
   
   With New FileSystemObject

      Select Case .FolderExists(FileSpec)
      
         Case Is = True
            
            ExistsFile = True
            
         Case Is = False
         
            If Create = True Then
                Call .CreateFolder(FileSpec)
                ExistsFile = True
            End If
            
      End Select
        
   End With
   
   CheckFolderExists = ExistsFile
     
End Function

Public Function CheckFileExists(ByVal FileSpec As String) As Boolean
    
    With New FileSystemObject
        
        CheckFileExists = .FileExists(FileSpec)
        
    End With
    
End Function


Public Function CheckIfUserYoutubeSubscribe() As Boolean
   
   With New FileSystemObject
      
      Dim TextFile As Variant
      Dim Line As String
      
      Set TextFile = .OpenTextFile(SysDirectorys.PathAppDef & "\Def.txt", ForReading, True)
      
        Do While Not TextFile.AtEndOfStream
            
            Line = TextFile.ReadLine
      
            If Line = "YoutubeSubscrib = Ok" Then
               CheckIfUserYoutubeSubscribe = True
               Exit Do
            End If
                
         Loop
      TextFile.Close
      
   End With
   
End Function

Public Function FindValueInString(ByVal StrX As String, ByVal StrY As String) As Boolean
    FindValueInString = VBA.InStr(1, StrX, StrY, vbTextCompare)
End Function

Public Function CheckConnectionProvider(ByVal Provider As String) As Integer

   On Error GoTo Exception
   
      Dim Database As ConnectionAccess
      
         Set Database = New ConnectionAccess
             Database.OpenConnection
             
             If Database.Connection.State = 1 Then
                CheckConnectionProvider = 1
             End If
             
         Set Database = Nothing
         
   Exit Function
   
Exception:

   CheckConnectionProvider = 0
   Set Database = Nothing
   
End Function

