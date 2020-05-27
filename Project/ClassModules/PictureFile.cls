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
         If CheckFile(FileString) = True Then
            Image.Picture = LoadPicture(FileString)
         End If
      Case Is <> Empty
         If CheckFile(OtherFile) = True Then
            Image.Picture = LoadPicture(OtherFile)
         End If
   End Select
End Sub

Public Sub CopyFile(Path As String, Dest As String)
   If CheckFile(Path) = True Then
      Dim File As FileSystemObject
         Set File = New FileSystemObject
            Call File.CopyFile(Path, Dest, True)
         Set File = Nothing
   End If
End Sub

Public Function DeleteFile(Path As String) As Boolean
   If CheckFile(Path) = True Then
      Call Kill(Path): DeleteFile = True
   End If
End Function

Public Function CheckFile(Path As String) As Boolean
   If Path = Empty Then Exit Function
   If Dir(Path, vbDirectory) <> vbNullString Then
      CheckFile = True
   End If
End Function

Public Property Get FileNumber() As String
   
   Dim Count As Integer
   Dim Number, ToString As String
   
   For Count = 1 To 25
      Number = Int((9 - 0 + 1) * Rnd + 0) & Chr(Int((22 * Rnd) + 1) + 64)
      ToString = ToString & Number
   Next
            
   FileNumber = ToString
                 
End Property

