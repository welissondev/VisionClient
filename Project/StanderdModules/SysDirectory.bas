Attribute VB_Name = "SysDirectory"
Option Explicit

Public Property Get PathSheet() As String
    PathSheet = ThisWorkbook.Path
End Property

Public Property Get PathClientPhoto() As String
    PathClientPhoto = ThisWorkbook.Path & "\User\Vision\ClientPhotos"
End Property

Public Property Get PathAppFileIcon() As String
    PathAppFileIcon = ThisWorkbook.Path & "\App\File\Icons"
End Property

Public Property Get PathAppLog() As String
    PathAppLog = ThisWorkbook.Path & "\App\Log"
End Property

Public Property Get PathAppDef() As String
   PathAppDef = ThisWorkbook.Path & "\App\Def"
End Property


