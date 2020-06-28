Attribute VB_Name = "SysDirectory"
Option Explicit

Public Property Get PathSheet() As String
    PathSheet = ThisWorkbook.Path
End Property

Public Property Get PathUserFileClientPhoto() As String
    PathUserFileClientPhoto = ThisWorkbook.Path & "\User\File\ClientPhoto"
End Property

Public Property Get PathUserDef() As String
    PathUserDef = ThisWorkbook.Path & "\User\Def"
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


