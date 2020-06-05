Attribute VB_Name = "AppSettings"
Option Explicit

'***************************************************
'Nesse módulo fica todos as propriedades que estão
'disponiveis publicamente para todo o sistema
'***************************************************

Public Property Get ConnectionString() As String
   ConnectionString = "Provider=Microsoft.ACE.Oledb.12.0;Data Source = " & ThisWorkbook.Path & "\App\Data\VisionBase.mdb"
End Property


Public Property Get ClientPhotosDirectory() As String
   ClientPhotosDirectory = ThisWorkbook.Path & "\User\Vision\ClientPhotos"
End Property
Public Property Get AppFileIconsDirectory() As String
   AppFileIconsDirectory = ThisWorkbook.Path & "\App\File\Icons"
End Property


Public Property Get AppVersion() As String
   AppVersion = "v1.2.14.160"
End Property
Public Property Get AppName() As String
   AppName = "Vision Client"
End Property


Public Property Get CompanyName() As String
   CompanyName = "Diário Excel"
End Property
Public Property Get CompanySite() As String
   CompanySite = "diarioexcel.com.br"
End Property
