Attribute VB_Name = "SysPropertys"
Option Explicit

Public Property Get ConnectionString() As String
   ConnectionString = "Provider = " & SheetAppUserDefination.BoxProviderSelected.Value & _
   "; Data Source = " & SysDirectorys.PathSheet & "\App\Data\VisionBase.mdb"
End Property

Public Property Get AppVersion() As String
   AppVersion = "1.3.30.1468"
End Property
Public Property Get AppName() As String
   AppName = "VisionClient"
End Property

Public Property Get CompanyName() As String
   CompanyName = "Diário Excel"
End Property
Public Property Get CompanySite() As String
   CompanySite = "diarioexcel.com.br"
End Property

Public Property Get YoutubeChannel()
   YoutubeChannel = "https://www.youtube.com/channel/UCSJAAxUzTj-qVVIKaqswQww?sub_confirmation=1"
End Property

