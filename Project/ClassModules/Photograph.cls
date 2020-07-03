VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Photograph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileString As String

Public Function GetFilePath() As String
    FileString = Application.GetOpenFilename("All Image files,*.jpg;*.jpg)")
    GetFilePath = FileString
End Function

Public Sub LoadFile(ByVal Image As MSForms.Image, Optional ByVal Path As String = Empty)
   Select Case Path
      Case Is = Empty
         If VerifyDirectoryFile(FileString) = True Then
            Image.Picture = LoadPicture(FileString)
         End If
      Case Is <> Empty
         If VerifyDirectoryFile(Path) = True Then
            Image.Picture = LoadPicture(Path)
         End If
   End Select
End Sub

Public Sub CopyFile(ByVal Sourc As String, ByVal Dest As String, Optional ByVal OverWiriteFile = True)
   If VerifyDirectoryFile(Sourc) = True Then
      With New FileSystemObject
         Call .CopyFile(Sourc, Dest, OverWiriteFile)
      End With
   End If
End Sub

Public Function DeleteFile(ByVal Path As String) As Boolean
   If VerifyDirectoryFile(Path) = True Then
      Call Kill(Path): DeleteFile = True
   End If
End Function

Public Function VerifyDirectoryFile(ByVal Path As String) As Boolean
   With New FileSystemObject
      VerifyDirectoryFile = .FileExists(Path)
   End With
End Function
