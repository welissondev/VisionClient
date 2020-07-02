VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
    Name As String
    Age As Integer
    Sex As String
    BirthDay As Date
    CivilStatus As String
    IndentyCard As String
    SocialSecurity As String
End Type


Private This As Class


Public Property Get Name() As String
    Name = This.Name
End Property
Public Property Let Name(Value As String)
    This.Name = Value
End Property


Public Property Get Age() As Integer
    Age = This.Age
End Property
Public Property Let Age(Value As Integer)
    This.Age = Value
End Property


Public Property Get Sex() As String
    Sex = This.Sex
End Property
Public Property Let Sex(Value As String)
    This.Sex = Value
End Property


Public Property Get BirthDay() As Date
    BirthDay = This.BirthDay
End Property
Public Property Let BirthDay(Value As Date)
    This.BirthDay = Value
End Property


Public Property Get CivilStatus() As String
    CivilStatus = This.CivilStatus
End Property
Public Property Let CivilStatus(Value As String)
    This.CivilStatus = Value
End Property


Public Property Get IndentyCard() As String
    IndentyCard = This.IndentyCard
End Property
Public Property Let IndentyCard(Value As String)
    This.IndentyCard = Value
End Property


Public Property Get SocialSecurity() As String
    SocialSecurity = This.SocialSecurity
End Property
Public Property Let SocialSecurity(Value As String)
    This.SocialSecurity = Value
End Property
