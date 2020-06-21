Attribute VB_Name = "SysProperty"
Option Explicit

'***************************************************
'Nesse módulo fica todos as propriedades que estão
'disponiveis publicamente para todo o sistema
'***************************************************

Public Property Get ConnectionString() As String
   ConnectionString = "Provider = " & SheetAppUserDefination.BoxProviderSelected.Text & _
   "; Data Source = " & ThisWorkbook.Path & "\App\Data\VisionBase.mdb"
End Property

Public Property Get ClientPhotosDirectory() As String
   ClientPhotosDirectory = ThisWorkbook.Path & "\User\Vision\ClientPhotos"
End Property
Public Property Get AppFileIconsDirectory() As String
   AppFileIconsDirectory = ThisWorkbook.Path & "\App\File\Icons"
End Property

Public Property Get AppVersion() As String
   AppVersion = "1.3.17.013"
End Property
Public Property Get AppName() As String
   AppName = "Sistema Vision Client"
End Property

Public Property Get CompanyName() As String
   CompanyName = "Diário Excel"
End Property
Public Property Get CompanySite() As String
   CompanySite = "diarioexcel.com.br"
End Property
