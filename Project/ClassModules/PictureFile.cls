VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PictureFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileString As String
    
Public Function OpenFile() As String
    FileString = Application.GetOpenFilename("All Image files,*.jpg;*.jpg)")
    OpenFile = FileString
End Function

Public Sub LoadFile(Image As MSForms.Image, Optional OtherFile As String = Empty)
   Select Case OtherFile
      Case Is = Empty
         If CheckFileDirectory(FileString) = True Then
            Image.Picture = LoadPicture(FileString)
         End If
      Case Is <> Empty
         If CheckFileDirectory(OtherFile) = True Then
            Image.Picture = LoadPicture(OtherFile)
         End If
   End Select
End Sub

Public Sub CopyFile(Path As String, Dest As String)
   If CheckFileDirectory(Path) = True Then
      Dim File As FileSystemObject
         Set File = New FileSystemObject
            Call File.CopyFile(Path, Dest, True)
         Set File = Nothing
   End If
End Sub

Public Function DeleteFile(Path As String) As Boolean
   If CheckFileDirectory(Path) = True Then
      Call Kill(Path): DeleteFile = True
   End If
End Function

Public Function CheckFileDirectory(Path As String) As Boolean
   If Path = Empty Then Exit Function
   If Dir(Path, vbDirectory) <> vbNullString Then
      CheckFileDirectory = True
   End If
End Function

