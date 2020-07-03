Attribute VB_Name = "SysDirectorys"
Option Explicit

Public Property Get PathSheet() As String
    PathSheet = ThisWorkbook.path
End Property

Public Property Get PathUserFileClientPhoto() As String
    PathUserFileClientPhoto = ThisWorkbook.path & "\User\File\ClientPhoto"
End Property

Public Property Get PathUserDef() As String
    PathUserDef = ThisWorkbook.path & "\User\Def"
End Property

Public Property Get PathAppFileIcon() As String
    PathAppFileIcon = ThisWorkbook.path & "\App\File\Icons"
End Property

Public Property Get PathAppLog() As String
    PathAppLog = ThisWorkbook.path & "\App\Log"
End Property

Public Property Get PathAppDef() As String
   PathAppDef = ThisWorkbook.path & "\App\Def"
End Property


