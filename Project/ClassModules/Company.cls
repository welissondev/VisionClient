VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
    Name As String
    FantasyName As String
    TimeDispatch As Integer
    YourType As String
    DateDispatch As Date
    TypeAction As String
    Ie As String
    Cnpj As String
End Type


Private This As Class


Public Property Get Name() As String
    Name = This.Name
End Property
Public Property Let Name(Value As String)
    This.Name = Value
End Property

Public Property Get FantasyName() As String
    FantasyName = This.FantasyName
End Property
Public Property Let FantasyName(Value As String)
    This.FantasyName = Value
End Property

Public Property Get TimeDispatch() As Integer
    TimeDispatch = This.TimeDispatch
End Property
Public Property Let TimeDispatch(Value As Integer)
    This.TimeDispatch = Value
End Property


Public Property Get YourType() As String
    YourType = This.YourType
End Property
Public Property Let YourType(Value As String)
    This.YourType = Value
End Property


Public Property Get DateDispatch() As Date
    DateDispatch = This.DateDispatch
End Property
Public Property Let DateDispatch(Value As Date)
    This.DateDispatch = Value
End Property


Public Property Get TypeAction() As String
    TypeAction = This.TypeAction
End Property
Public Property Let TypeAction(Value As String)
    This.TypeAction = Value
End Property


Public Property Get Ie() As String
    Ie = This.Ie
End Property
Public Property Let Ie(Value As String)
    This.Ie = Value
End Property


Public Property Get Cnpj() As String
    Cnpj = This.Cnpj
End Property
Public Property Let Cnpj(Value As String)
    This.Cnpj = Value
End Property

