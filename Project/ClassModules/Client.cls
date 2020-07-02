VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
  Id As Long
  PhotoNumber As String
  ActiveStatus As String
  InternalCode As String
  RegistrationDate As Date
  Observation As String
End Type


Private This As Class


Public Property Get Id() As Long
    Id = This.Id
End Property
Public Property Let Id(Value As Long)
    This.Id = Value
End Property

Public Property Get PhotoNumber() As String
    PhotoNumber = This.PhotoNumber
End Property
Public Property Let PhotoNumber(Value As String)
    This.PhotoNumber = Value
End Property

Public Property Get ActiveStatus() As String
    ActiveStatus = This.ActiveStatus
End Property
Public Property Let ActiveStatus(Value As String)
    This.ActiveStatus = Value
End Property

Public Property Get InternalCode() As String
    InternalCode = This.InternalCode
End Property
Public Property Let InternalCode(Value As String)
    This.InternalCode = Value
End Property

Public Property Get RegistrationDate() As Date
    RegistrationDate = This.RegistrationDate
End Property
Public Property Let RegistrationDate(Value As Date)
    This.RegistrationDate = Value
End Property

Public Property Get Observation() As String
    Observation = This.Observation
End Property
Public Property Let Observation(Value As String)
    This.Observation = Value
End Property

