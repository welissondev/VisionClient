VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhysicalCustomerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private This As Class

Private Type Class
  Id As Integer
  BirthDay As Date
  Age As Integer
  State As String
  CivilStatus As String
  Sex As String
  PhotoNumber As String
End Type

Public InternalCode As String
Public YourName As String
Public IndentyCard As String
Public SocialSecurity As String

Public FixedPhone As String
Public MobilePhone As String
Public WhatsApp As String
Public Email As String

Public District As String
Public City As String
Public ZipCode As String
Public StreetNumber As String
Public AddressDescription As String
Public AddressComplement As String
Public AddressNote As String

Public ActiveStatus As Boolean
Public RegistrationDate As Date


Public Function Insert() As Boolean
    
   Dim Database As ConnectionAccessModel
   Dim QueryBuilder As StringBuilder
   Dim Sql As String
    
   Set QueryBuilder = New StringBuilder
   
      With QueryBuilder
         .Append "INSERT INTO physicalCustomer(internalCode, yourName, age, birthDay, sex, indentyCard, socialSecurity, civilStatus, photoNumber, "
         .Append "fixedPhone, mobilePhone, whatSapp, email, district, city, state, zipCode, streetNumber, addressDescription, addressComplement, "
         .Append "addressNote, registrationDate, ActiveStatus) VALUES (@internalCode, @yourName, @age, @birthDay, @sex, @indentyCard,"
         .Append "@socialSecurity, @civilStatus, @photoNumber,@fixedPhone, @mobilePhone, @whatSapp, @email, @district, @city, @state,"
         .Append "@zipCode, @streetNumber, @addressDescription, @addressComplement, @addressNote, @registrationDate, @ActiveStatus);"
      End With
      Sql = QueryBuilder.ToString
         
   Set Database = New ConnectionAccessModel
      With Database
         .AddParameter "@internalCode", InternalCode, adVarChar
         .AddParameter "@yourName", YourName, adVarChar
         .AddParameter "@age", This.Age, adInteger
         .AddParameter "@birthDay", This.BirthDay, adDate
         .AddParameter "@sex", This.Sex, adVarChar
         .AddParameter "@indentyCard", IndentyCard, adVarChar
         .AddParameter "@socialSecurity", SocialSecurity, adVarChar
         .AddParameter "@civilStatus", This.CivilStatus, adVarChar
         .AddParameter "@photoNumber", This.PhotoNumber, adVarChar
         .AddParameter "@fixedPhone", FixedPhone, adVarChar
         .AddParameter "@mobilePhone", MobilePhone, adVarChar
         .AddParameter "@whatSapp", WhatsApp, adVarChar
         .AddParameter "@email", Email, adVarChar
         .AddParameter "@district", District, adVarChar
         .AddParameter "@City", City, adVarChar
         .AddParameter "@state", This.State, adVarChar
         .AddParameter "@zipCode", ZipCode, adVarChar
         .AddParameter "@streetNumber", StreetNumber, adVarChar
         .AddParameter "@addressDescription", AddressDescription, adVarChar
         .AddParameter "@addressComplement", AddressComplement, adVarChar
         .AddParameter "@addressNote", AddressNote, adVarChar
         .AddParameter "@registrationDate", Date, adDate
         .AddParameter "@ActiveStatus", ActiveStatus, adBoolean
      End With
      
      Insert = Database.ExecuteNonQuery(Sql)
   
   Set Database = Nothing
   Set QueryBuilder = Nothing
    
End Function

Public Function Update(Id As Integer) As Boolean

   Dim Database As ConnectionAccessModel
   Dim QueryBilder As StringBuilder
   Dim Sql As String

   Set QueryBilder = New StringBuilder
      With QueryBilder
         .Append "UPDATE PhysicalCustomer Set internalCode = @internalCode, yourName = @yourName, age = @age,"
         .Append "birthDay = @birthDay, sex = @sex, indentyCard = @indentyCard, socialSecurity = @socialSecurity,"
         .Append "civilStatus = @civilStatus, fixedPhone = @fixedPhone, mobilePhone = @mobilePhone, whatSapp = @whatSapp,"
         .Append "email = @email, addressDescription = @AddressDescription, district = @district, city = @city,"
         .Append "state = @state, zipCode = @zipCode, streetNumber = @streetNumber, addressComplement = @addressComplement,"
         .Append "ActiveStatus = @ActiveStatus, addressNote = @addressNote, photoNumber = @photoNumber,"
         .Append "registrationDate = @registrationDate WHERE id = @id"
      End With
      Sql = QueryBilder.ToString
                                             
   Set Database = New ConnectionAccessModel
      With Database
         .AddParameter "@internalCode", InternalCode, adVarChar
         .AddParameter "@yourName", YourName, adVarChar
         .AddParameter "@age", This.Age, adInteger
         .AddParameter "@birthDay", This.BirthDay, adDate
         .AddParameter "@sex", This.Sex, adVarChar
         .AddParameter "@indentyCard", IndentyCard, adVarChar
         .AddParameter "@socialSecurity", SocialSecurity, adVarChar
         .AddParameter "@civilStatus", This.CivilStatus, adVarChar
         .AddParameter "@fixedPhone", FixedPhone, adVarChar
         .AddParameter "@mobilePhone", MobilePhone, adVarChar
         .AddParameter "@whatSapp", WhatsApp, adVarChar
         .AddParameter "@email", Email, adVarChar
         .AddParameter "@addressDescription", AddressDescription, adVarChar
         .AddParameter "@district", District, adVarChar
         .AddParameter "@city", City, adVarChar
         .AddParameter "@state", This.State, adVarChar
         .AddParameter "@zipCode", ZipCode, adVarChar
         .AddParameter "@streetNumber", StreetNumber, adVarChar
         .AddParameter "@addressComplement", AddressComplement, adVarChar
         .AddParameter "@ActiveStatus", ActiveStatus, adBoolean
         .AddParameter "@addressNote", AddressNote, adVarChar
         .AddParameter "@photoNumber", PhotoNumber, adVarChar
         .AddParameter "@registrationDate", RegistrationDate, adDate
         .AddParameter "@id", Id, adInteger
      End With
      
      Update = Database.ExecuteNonQuery(Sql)
   
   Set Database = Nothing
   Set QueryBilder = Nothing

End Function

Public Function Delete(Coll As Collection) As Boolean

   Dim Database As ConnectionAccessModel
   Dim QueryBuilder As StringBuilder
   Dim Sql As String
   Dim Index As Integer
   
   If MsgBox("Confirma a exclusão desse cliente?", _
   vbExclamation + vbYesNo + vbDefaultButton2, "IMPORTANTE") = _
   vbNo Then Exit Function
      
   Set QueryBuilder = New StringBuilder
      With QueryBuilder
         .Append "DELETE FROM physicalCustomer WHERE id = @id"
      End With
      Sql = QueryBuilder.ToString
      
   Set Database = New ConnectionAccessModel
      For Index = 1 To Coll.Count
         With Database
            .ClearParameter
            .AddParameter "@id", Coll(Index), adInteger
             Delete = .ExecuteNonQuery(Sql)
         End With
      Next
      
   Set Database = Nothing
   Set QueryBuilder = Nothing
   
End Function

Public Sub GetDetails(Id As Integer)

   Dim Database As ConnectionAccessModel
   Dim RecordSet As ADODB.RecordSet
   Dim QueryBuilder As StringBuilder
   Dim Sql As String

   Set QueryBuilder = New StringBuilder
      With QueryBuilder
         .Append "SELECT * FROM PhysicalCustomer WHERE Id = @Id"
      End With

      Sql = QueryBuilder.ToString

   Set Database = New ConnectionAccessModel
      With Database
         .AddParameter "@internalCode", Id, adVarChar
         Set RecordSet = .ExecuteWithQuery(Sql)
      End With

      If RecordSet.BOF = False Then
         With RecordSet
            This.Id = .Fields("id").Value
            InternalCode = .Fields("internalCode").Value
            YourName = .Fields("yourName").Value
            Age = .Fields("age").Value
            BirthDay = .Fields("birthDay").Value
            Sex = .Fields("sex").Value
            IndentyCard = .Fields("indentyCard").Value
            SocialSecurity = .Fields("socialSecurity").Value
            CivilStatus = .Fields("civilStatus").Value
            FixedPhone = .Fields("fixedPhone").Value
            MobilePhone = .Fields("mobilePhone").Value
            WhatsApp = .Fields("whatsApp").Value
            Email = .Fields("email").Value
            AddressDescription = .Fields("addressDescription").Value
            District = .Fields("district").Value
            City = .Fields("city").Value
            State = .Fields("state").Value
            ZipCode = .Fields("zipCode").Value
            StreetNumber = .Fields("streetNumber").Value
            AddressComplement = .Fields("addressComplement").Value
            ActiveStatus = .Fields("ActiveStatus").Value
            AddressNote = .Fields("addressNote").Value
            PhotoNumber = .Fields("photoNumber").Value
            RegistrationDate = .Fields("registrationDate").Value
         End With
      End If

   Set QueryBuilder = Nothing
   Set Database = Nothing

End Sub

Public Property Get Id() As Variant
   Id = This.Id
End Property
Public Property Let Id(Value As Variant)
   This.Id = Value
End Property

Public Property Get BirthDay() As Variant
   If This.BirthDay <> "00:00:00" Then
      BirthDay = This.BirthDay
   End If
End Property
Public Property Let BirthDay(Value As Variant)
   If Value <> vbNullString And IsNull(Value) = False And IsDate(Value) = True Then
      This.BirthDay = Value
   End If
End Property

Public Property Get Age() As Variant
   If This.Age > 0 Then Age = This.Age
End Property
Public Property Let Age(Value As Variant)
   If Value <> vbNullString And IsEmpty(Value) = False And IsNumeric(Value) = True Then
      This.Age = Value
   End If
End Property

Public Property Get Sex() As String
   If This.Sex = Empty Then
      Sex = "Select"
   Else
      Sex = This.Sex
   End If
End Property
Public Property Let Sex(Value As String)
   If Value = "Select" Then
      This.Sex = Empty
   Else
      This.Sex = Value
   End If
End Property

Public Property Get CivilStatus() As String
   If This.CivilStatus = Empty Then
      CivilStatus = "Select"
   Else
      CivilStatus = This.CivilStatus
   End If
End Property
Public Property Let CivilStatus(Value As String)
   If Value = "Select" Then
      This.CivilStatus = Empty
   Else
      This.CivilStatus = Value
   End If
End Property

Public Property Get State() As String
   If This.State = Empty Then
      State = "Select"
   Else
      State = This.State
   End If
End Property
Public Property Let State(Value As String)
   If Value = "Select" Then
      This.State = Empty
   Else
      This.State = Value
   End If
End Property

Public Property Get PhotoNumber() As Variant
   PhotoNumber = This.PhotoNumber
End Property
Public Property Let PhotoNumber(Value As Variant)
   If IsNull(Value) = True Then
      This.PhotoNumber = Empty
   Else
      This.PhotoNumber = Value
   End If
End Property
