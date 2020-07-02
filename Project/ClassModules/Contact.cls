VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
    FixedPhone As String
    MobilePhone As String
    WhatsApp As String
    Email As String
End Type


Private This As Class


Public Property Get FixedPhone() As String
    FixedPhone = This.FixedPhone
End Property
Public Property Let FixedPhone(Value As String)
    This.FixedPhone = Value
End Property


Public Property Get MobilePhone() As String
    MobilePhone = This.MobilePhone
End Property
Public Property Let MobilePhone(Value As String)
    This.MobilePhone = Value
End Property


Public Property Get WhatsApp() As String
    WhatsApp = This.WhatsApp
End Property
Public Property Let WhatsApp(Value As String)
    This.WhatsApp = Value
End Property


Public Property Get Email() As String
    Email = This.Email
End Property
Public Property Let Email(Value As String)
    This.Email = Value
End Property
