VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClientNaturalPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Client As Client
Private Person As Person
Private Address As Address
Private Contact As Contact


Public Sub Builder(Object As Object)
    Set Client = Object
    Set Person = Object
    Set Address = Object
    Set Contact = Object
End Sub


Public Function Insert() As Boolean
    
    Dim Base As ConnectionAccess
    Dim Query As StringBuilder

    Set Base = New ConnectionAccess
    Set Query = New StringBuilder
        
    With Query
       .Append "INSERT INTO ClientNaturalPerson(internalCode, photoNumber, activeStatus, registrationDate, observation,"
       .Append "yourName, age, sex, birthDay, civilStatus, socialSecurity, indentyCard, fixedPhone, mobilePhone,"
       .Append "whatSapp, email, district, city, state, zipCode, streetNumber, addressDescription, addressComplement)"
       .Append "VALUES(@internalCode, @photoNumber, @activeStatus, @registrationDate, @observation, @yourName, @age, @sex,"
       .Append "@birthDay, @civilStatus, @socialSecurity, @indentyCard, @fixedPhone, @mobilePhone, @whatsApp, @email,"
       .Append "@district, @city, @state, @zipCode, @streetNumber, @addressDescription, @addressComplement)"
    End With
                   
    With Client
        Base.AddParameter "@internalCode", .InternalCode, adVarChar
        Base.AddParameter "@photoNumber", .PhotoNumber, adVarChar
        Base.AddParameter "@activeStatus", .ActiveStatus, adVarChar
        Base.AddParameter "@registrationDate", Now(), adDate
        Base.AddParameter "@observation", .Observation, adVarChar
    End With
         
    With Person
        Base.AddParameter "@yourName", .Name, adVarChar
        Base.AddParameter "@age", .Age, adNumeric
        Base.AddParameter "@sex", .Sex, adVarChar
        Base.AddParameter "@birthDay", .BirthDay, adDate
        Base.AddParameter "@civilStatus", .CivilStatus, adVarChar
        Base.AddParameter "@socialSecurity", .SocialSecurity, adVarChar
        Base.AddParameter "@indentyCard", .IndentyCard, adVarChar
    End With
                                     
    With Contact
        Base.AddParameter "@fixedPhone", .FixedPhone, adVarChar
        Base.AddParameter "@mobilePhone", .MobilePhone, adVarChar
        Base.AddParameter "@whatsApp", .WhatsApp, adVarChar
        Base.AddParameter "@email", .Email, adVarChar
    End With
                                     
    With Address
        Base.AddParameter "@district", .District, adVarChar
        Base.AddParameter "@city", .City, adVarChar
        Base.AddParameter "@state", .State, adVarChar
        Base.AddParameter "@zipCode", .ZipCode, adVarChar
        Base.AddParameter "@streetNumber", .StreetNumber, adVarChar
        Base.AddParameter "@addressDescription", .Description, adVarChar
        Base.AddParameter "@addressComplement", .Complement, adVarChar
    End With
  
    Insert = Base.ExecuteNonQuery(Query.ToString())
    
   
   Set Base = Nothing
   Set Query = Nothing
    
End Function

Public Function Update() As Boolean

      Dim Base As ConnectionAccess
      Dim Query As StringBuilder
  
      Set Query = New StringBuilder
      Set Base = New ConnectionAccess
   
      With Query
        .Append "UPDATE ClientNaturalPerson Set internalCode = @internalCode, photoNumber = @photoNumber, activeStatus = @activeStatus,"
        .Append "registrationDate = @registrationDate, observation  = @observation, yourName = @yourName, age = @age, sex = @sex,"
        .Append "birthDay = @birthDay, civilStatus = @civilStatus, socialSecurity = @socialSecurity, indentyCard = @indentyCard,"
        .Append "fixedPhone = @fixedPhone, mobilePhone = @mobilePhone, whatsApp = @whatsApp, email = @email, district = @district,"
        .Append "city = @city, state = @state, zipCode = @zipCode, streetNumber = @streetNumber, addressDescription = @addressDescription,"
        .Append "addressComplement = @addressComplement WHERE id = @id"
      End With
      
      With Client
        Base.AddParameter "@internalCode", .InternalCode, adVarChar
        Base.AddParameter "@photoNumber", .PhotoNumber, adVarChar
        Base.AddParameter "@activeStatus", .ActiveStatus, adVarChar
        Base.AddParameter "@registrationDate", Now(), adDate
        Base.AddParameter "@observation", .Observation, adVarChar
      End With
           
      With Person
        Base.AddParameter "@yourName", .Name, adVarChar
        Base.AddParameter "@age", .Age, adNumeric
        Base.AddParameter "@sex", .Sex, adVarChar
        Base.AddParameter "@birthDay", .BirthDay, adDate
        Base.AddParameter "@civilStatus", .CivilStatus, adVarChar
        Base.AddParameter "@socialSecurity", .SocialSecurity, adVarChar
        Base.AddParameter "@indentyCard", .IndentyCard, adVarChar
      End With
                                     
      With Contact
        Base.AddParameter "@fixedPhone", .FixedPhone, adVarChar
        Base.AddParameter "@mobilePhone", .MobilePhone, adVarChar
        Base.AddParameter "@whatsApp", .WhatsApp, adVarChar
        Base.AddParameter "@email", .Email, adVarChar
      End With
                                     
      With Address
        Base.AddParameter "district", .District, adVarChar
        Base.AddParameter "city", .City, adVarChar
        Base.AddParameter "state", .State, adVarChar
        Base.AddParameter "zipCode", .ZipCode, adVarChar
        Base.AddParameter "streetNumber", .StreetNumber, adVarChar
        Base.AddParameter "addressDescription", .Description, adVarChar
        Base.AddParameter "addressComplement", .Complement, adVarChar
      End With
      
      With Client
        Base.AddParameter "@id", .Id, adNumeric
      End With
      
      Update = Base.ExecuteNonQuery(Query.ToString())
   
   Set Base = Nothing
   Set Query = Nothing

End Function

Public Function Delete(Id As Long) As Boolean

    Dim Base As ConnectionAccess
    Dim Query As StringBuilder
    
    If MsgBox("Confirma a exclusão desse cliente?", _
    vbExclamation + vbYesNo + vbDefaultButton2, "IMPORTANTE") = _
    vbNo Then Exit Function
    
    Set Base = New ConnectionAccess
    Set Query = New StringBuilder
    
    With Query
      .Append "DELETE FROM ClientNaturalPerson WHERE id = @id"
    End With
  
    With Base
       .AddParameter "@id", Id, adNumeric
    End With
    
    Delete = Base.ExecuteNonQuery(Query.ToString())
     
    Set Base = Nothing
    Set Query = Nothing
   
End Function


Public Sub ViewData(Id As Long)

    Dim Base As ConnectionAccess
    Dim Query As StringBuilder
    
    Set Base = New ConnectionAccess
    Set Query = New StringBuilder
    
    With Query
       .Append "SELECT * FROM ClientNaturalPerson WHERE Id = @Id"
    End With
        
    With Base
       .AddParameter "@Id", Id, adNumeric
    End With
    
    Call Base.ExecuteWithQuery(Query.ToString())
      
    If Base.Recordset.EOF = False Then
       
       With Client
          .Id = Base.Field("id")
          .InternalCode = Base.Field("internalCode")
          .PhotoNumber = Base.Field("photoNumber")
          .RegistrationDate = Base.Field("registrationDate")
          .ActiveStatus = Base.Field("ActiveStatus")
          .Observation = Base.Field("observation")
       End With
       
       With Person
          .Name = Base.Field("yourName")
          .Age = Base.Field("age")
          .Sex = Base.Field("sex")
          .BirthDay = Base.Field("birthDay")
          .CivilStatus = Base.Field("civilStatus")
          .IndentyCard = Base.Field("indentyCard")
          .SocialSecurity = Base.Field("socialSecurity")
       End With
       
       With Contact
          .FixedPhone = Base.Field("fixedPhone")
          .MobilePhone = Base.Field("mobilePhone")
          .WhatsApp = Base.Field("whatsApp")
          .Email = Base.Field("email")
       End With
       
       With Address
          .District = Base.Field("district")
          .City = Base.Field("city")
          .State = Base.Field("state")
          .ZipCode = Base.Field("zipCode")
          .StreetNumber = Base.Field("streetNumber")
          .Description = Base.Field("addressDescription")
          .Complement = Base.Field("addressComplement")
       End With
       
    End If
    
    Set Base = Nothing
    Set Query = Nothing

End Sub

Public Function ListData() As ADODB.Recordset
   
    Dim Base As ConnectionAccess
    Dim Query As StringBuilder
    
    Set Base = New ConnectionAccess
    Set Query = New StringBuilder
    
    With Query
       .Append "SELECT id, internalCode, yourName, socialSecurity, indentyCard,"
       .Append "email, sex, civilStatus, activeStatus FROM "
       .Append "ClientNaturalPerson WHERE Id > 0 "
    End With
              
    Set ListData = Base.ExecuteWithQuery(Query.ToString())
         
    Set Base = Nothing
    Set Query = Nothing

End Function




