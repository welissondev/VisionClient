VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
    District As String
    City As String
    State As String
    ZipCode As String
    StreetNumber As String
    Description As String
    Complement As String
End Type


Private This As Class


Public Property Get District() As String
    District = This.District
End Property
Public Property Let District(Value As String)
    This.District = Value
End Property


Public Property Get City() As String
    City = This.City
End Property
Public Property Let City(Value As String)
    This.City = Value
End Property


Public Property Get State() As String
    State = This.State
End Property
Public Property Let State(Value As String)
    This.State = Value
End Property


Public Property Get ZipCode() As String
    ZipCode = This.ZipCode
End Property
Public Property Let ZipCode(Value As String)
    This.ZipCode = Value
End Property


Public Property Get StreetNumber() As String
    StreetNumber = This.StreetNumber
End Property
Public Property Let StreetNumber(Value As String)
    This.StreetNumber = Value
End Property


Public Property Get Description() As String
    Description = This.Description
End Property
Public Property Let Description(Value As String)
    This.Description = Value
End Property


Public Property Get Complement() As String
    Complement = This.Complement
End Property
Public Property Let Complement(Value As String)
    This.Complement = Value
End Property


