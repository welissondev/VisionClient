Attribute VB_Name = "AppConfig"
Option Explicit

Public Property Get ConnectionStrig() As String
   ConnectionStrig = "Provider=Microsoft.ACE.Oledb.12.0;Data Source = " & ThisWorkbook.Path & "\App\Data\VisionBase.mdb"
End Property

Public Property Get ClientPhotosDirectory() As String
   ClientPhotosDirectory = ThisWorkbook.Path & "\User\Vision\ClientPhotos"
End Property

Public Property Get AppFileIconsDirectory() As String
   AppFileIconsDirectory = ThisWorkbook.Path & "\App\File\Icons"
End Property
