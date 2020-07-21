VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAddClientNaturalPerson 
   Caption         =   "Formulário Para Cadastro De Clientes Tipo Física"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   OleObjectBlob   =   "FormAddClientNaturalPerson.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FormAddClientNaturalPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Id As Long
Private FileString As String
Private PhotoNumber As String
Private RegistrationDate As Date
Private Mask() As New FormatterMask

Private Sub UserForm_Initialize()

   Call FormatMask
   Call FillTextBox
   Call SysMethods.DefineUserFormStyle(Me, 11892015)
   Call LoadImageNothing(ImageClient)

End Sub

Private Sub ButtonSelectPhoto_Click()

   On Error GoTo Exception

      Dim Image As MSForms.Image
      Dim Picture As Photograph
      Dim Number As CodeGenerator
      Dim Path As String
      Dim Directory As Boolean

      Set Image = ImageClient
      Set Picture = New Photograph
      Set Number = New CodeGenerator

      Path = Picture.GetFilePath()
      Directory = Picture.VerifyDirectoryFile(Path)

      Select Case Directory
         Case Is = True

            If Picture.VerifyDirectoryFile(CurrentPhoto) = False Then
               PhotoNumber = "PF-" & Number.Generate(25, True)
            End If

            FileString = Picture.FileString

            Call Picture.LoadFile(Image)

         Case Is = False

            If Picture.VerifyDirectoryFile(CurrentPhoto) = False Then
               Call LoadImageNothing(Image)
            Else
               Call Picture.LoadFile(Image, CurrentPhoto)
            End If

      End Select

      Set Picture = Nothing
      Set Number = Nothing

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub


Private Sub ButtonSave_Click()

   On Error GoTo Exception
        
      If TextActiveStatus.Text = "Selecionar" Then
            MsgBox "Defina o estado de atividade do cliente!", vbExclamation, "Obrigatório"
            TextActiveStatus.SetFocus
          Exit Sub
      End If
          
      If TextInternalCode.Text = Empty Then
           MsgBox "Campo código é obrigatório!", vbExclamation, "OBRIGATÓRIO"
           TextInternalCode.SetFocus
          Exit Sub
      End If
        
      If Len(TextInternalCode.Value) > 8 Then
            MsgBox "O código do cliente, deve ter no máximo 8 digitos!", vbExclamation, "Obrigatório"
            TextInternalCode.SetFocus
          Exit Sub
      End If

      If TextYourName.Text = Empty Then
            MsgBox "O nome do cliente não foi informado!", vbExclamation, "Obrigatório"
            TextYourName.SetFocus
          Exit Sub
      End If
      
      If TextAge.Text = Empty Or TextAge.Value = 0 Then
            MsgBox "Idade não pode ser zero ou vazia!", vbExclamation, "Obrigatório"
            TextAge.SetFocus
          Exit Sub
      End If
      
      If IsDate(TextBirthDay.Value) = False Then
            MsgBox "A data de nascimento não é válida!", vbExclamation, "INVÁLIDA"
            TextBirthDay.SetFocus
          Exit Sub
      End If
      
      Select Case Id
      
         Case Is = 0
         
            If Insert() = True Then
                
                With New Photograph
                    Call .CopyFile(FileString, CurrentPhoto)
                End With
                
               Call ButtonClear_Click
               
               MsgBox "Registrado com sucesso!", _
               vbInformation, "SUCESSO"
            
            End If
            
         Case Is > 0
         
            If Update(Id) = True Then
                
              With New Photograph
                  Call .CopyFile(FileString, CurrentPhoto)
              End With
            
              MsgBox "Editado com sucesso!", _
              vbInformation, "SUCESSO"
            
            End If
            
      End Select

  Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub ButtonClear_Click()

   On Error GoTo Exception

      Dim Control As Control

      For Each Control In Me.Controls

         Select Case TypeName(Control)

            Case Is = "TextBox"
               Control.Value = Empty

            Case Is = "ComboBox"
               Control.Text = Control.List(0)

            Case Is = "CheckBox"
               Control.Value = False

            Case Is = "Image"
               Call LoadImageNothing(Control)

         End Select
      Next

      With TextInternalCode
         .Locked = False
      End With

      With TextActiveStatus
         .SetFocus
      End With

      Id = 0
      FileString = Empty
      PhotoNumber = Empty

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub ButtonDelete_Click()

   On Error GoTo Exception

        If Delete(Id) = True Then
            
            With New Photograph
                Call .DeleteFile(CurrentPhoto)
            End With
  
           Call ButtonClear_Click
  
           MsgBox "Deletado com sucesso!", _
           vbInformation, "Sucesso"
            
        End If

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub ButtonClose_Click()
   Unload Me
End Sub

Private Sub CheckGenerateCode_Click()

   On Error GoTo Exception

      Dim Control As MSForms.CheckBox
      Dim Code As CodeGenerator

      Set Control = CheckGenerateCode
      Set Code = New CodeGenerator

      Select Case Control.Value
         Case Is = True
            With TextInternalCode
               .Text = Code.Generate(4, True)
               .Locked = True
               .SetFocus
            End With
         Case Is = False
            With TextInternalCode
               .Text = Empty
               .Locked = False
               .SetFocus
            End With
      End Select

      Set Code = Nothing

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub LoadImageNothing(Image As MSForms.Image)

   On Error GoTo Exception

      With New Photograph
         Call .LoadFile(Image, ImageNothing)
      End With

      FileString = Empty
      PhotoNumber = Empty

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub FillTextBox()

   On Error GoTo Exception
      Call SysCollections.SetYesNo(TextActiveStatus)
      Call SysCollections.SetSexes(TextSex)
      Call SysCollections.SetCivilStatus(TextCivilStatus)
      Call SysCollections.SetStatesLocation(TextState)
   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Sub FormatMask()

   On Error GoTo Exception

      Dim Count As Integer
      Dim Index As Integer

      Count = Me.Controls.Count - 1

      ReDim Mask(0 To Count)

      For Index = 0 To Count
         Select Case Me.Controls(Index).Tag
            Case Is = "Date"
                Set Mask(Index).ToDate = Controls(Index)
            Case Is = "MobilePhone"
                Set Mask(Index).ToMobilePhone = Controls(Index)
            Case Is = "FixedPhone"
                Set Mask(Index).ToFixedPhone = Controls(Index)
            Case Is = "SocialSecurity"
                Set Mask(Index).ToSocialSecurity = Controls(Index)
            Case Is = "ZipCode"
                Set Mask(Index).ToZipCode = Controls(Index)
            Case Is = "Number"
                Set Mask(Index).CanNotString = Controls(Index)
         End Select
      Next

   Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub


Private Function ValidateToInsert() As Boolean
    
    Dim Query As StringBuilder
    Set Query = New StringBuilder
    
    With New ConnectionAccess
    
        '\\Valida se existe cliente cadastrado com o código
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE internalCode = @internalCode"
        .AddParameter "@internalCode", TextInternalCode.Text, adVarChar
        
        If .ExecuteWithQuery(Query.ToString()).RecordCount > 0 Then
               MsgBox "Um cliente cadastrado com esse código já existe!", _
               vbExclamation, "JÁ EXISTE"
               TextInternalCode.SetFocus
               ValidateToInsert = False
            Exit Function
        End If
        Query.Clear: .ClearParameter
                     
        '\\Valida se existe cliente cadastrado com o CPF
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE socialSecurity = @socialSecutiry"
        .AddParameter "@socialSecutiry", TextSocialSecurity.Text, adVarChar
        
        If .ExecuteWithQuery(Query.ToString()).RecordCount > 0 Then
              MsgBox "Um cliente cadastrado com esse CPF já existe!", _
              vbExclamation, "JÁ EXISTE"
              TextSocialSecurity.SetFocus
              ValidateToInsert = False
            Exit Function
        End If
        Query.Clear: .ClearParameter
             
             
        '\\Valida se existe cliente cadastrado com o RG
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE indentyCard = @indentyCard"
        .AddParameter "@indentyCard", TextIndentyCard.Text, adVarChar
        
        If .ExecuteWithQuery(Query.ToString()).RecordCount > 0 Then
              MsgBox "Um cliente cadastrado com esse RG já existe!", _
              vbExclamation, "JÁ EXISTE"
              TextIndentyCard.SetFocus
              ValidateToInsert = False
            Exit Function
        End If
        Query.Clear: .ClearParameter
        
        
        '\\Valida se existe cliente cadastrado com o email
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE email = @email"
        .AddParameter "@email", TextEmail.Text, adVarChar
        
        If .ExecuteWithQuery(Query.ToString()).RecordCount > 0 Then
              MsgBox "Um cliente cadastrado com esse email já existe!", _
              vbExclamation, "JÁ EXISTE"
              TextEmail.SetFocus
              ValidateToInsert = False
            Exit Function
        End If
        Query.Clear: .ClearParameter
             
        ValidateToInsert = True
        
    End With
End Function

Private Function Insert() As Boolean

    If ValidateToInsert() = False Then Exit Function

    Dim Query As StringBuilder
    Set Query = New StringBuilder
     
    With New ConnectionAccess
    
        With Query
           .Append "INSERT INTO ClientNaturalPerson(internalCode, photoNumber, activeStatus, registrationDate, observation,"
           .Append "yourName, age, sex, birthDay, civilStatus, socialSecurity, indentyCard, fixedPhone, mobilePhone,"
           .Append "whatSapp, email, district, city, state, zipCode, streetNumber, addressDescription, addressComplement)"
           .Append "VALUES(@internalCode, @photoNumber, @activeStatus, @registrationDate, @observation, @yourName, @age, @sex,"
           .Append "@birthDay, @civilStatus, @socialSecurity, @indentyCard, @fixedPhone, @mobilePhone, @whatsApp, @email,"
           .Append "@district, @city, @state, @zipCode, @streetNumber, @addressDescription, @addressComplement)"
        End With
      
        '\\Cliente
          .AddParameter "@internalCode", TextInternalCode.Text, adVarChar
          .AddParameter "@photoNumber", PhotoNumber, adVarChar
          .AddParameter "@activeStatus", TextActiveStatus.Text, adVarChar
          .AddParameter "@registrationDate", Date, adDate
          .AddParameter "@observation", TextObservation.Text, adVarChar
          .AddParameter "@yourName", TextYourName.Text, adVarChar
          .AddParameter "@age", TextAge.Text, adNumeric
          .AddParameter "@sex", TextSex.Text, adVarChar
          .AddParameter "@birthDay", TextBirthDay.Text, adDate
          .AddParameter "@civilStatus", TextCivilStatus.Text, adVarChar
          .AddParameter "@socialSecurity", TextSocialSecurity.Text, adVarChar
          .AddParameter "@indentyCard", TextIndentyCard.Text, adVarChar
                                       
        '\\Contato
          .AddParameter "@fixedPhone", TextFixedPhone.Text, adVarChar
          .AddParameter "@mobilePhone", TextMobilePhone.Text, adVarChar
          .AddParameter "@whatsApp", TextWhatsapp.Text, adVarChar
          .AddParameter "@email", TextEmail.Text, adVarChar
                                       
        '\\Endereço
          .AddParameter "district", TextDistrict.Text, adVarChar
          .AddParameter "city", TextCity.Text, adVarChar
          .AddParameter "state", TextState.Text, adVarChar
          .AddParameter "zipCode", TextZipCode.Text, adVarChar
          .AddParameter "streetNumber", TextStreetNumber.Text, adVarChar
          .AddParameter "addressDescription", TextAddressDescription.Text, adVarChar
          .AddParameter "addressComplement", TextAddressComplement.Text, adVarChar
    
          Insert = .ExecuteNonQuery(Query.ToString())
    
    End With
   
    Set Query = Nothing

End Function


Private Function ValidateToUpdate() As Boolean
    
    Dim Query As StringBuilder
    Set Query = New StringBuilder
    
    With New ConnectionAccess
    
        '\\Valida se existe cliente cadastrado com o código
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE internalCode = @internalCode"
        .AddParameter "@internalCode", TextInternalCode.Text, adVarChar
        
        Call .ExecuteWithQuery(Query.ToString())
        
        If Not .RecordSet.EOF Then
            If .RecordSet.Fields("Id").Value > 0 And .RecordSet.Fields("Id").Value <> Id Then
                   MsgBox "Um cliente cadastrado com esse código já existe!", _
                   vbExclamation, "JÁ EXISTE"
                   TextInternalCode.SetFocus
                   ValidateToUpdate = False
                Exit Function
            End If
        End If
        Query.Clear: .ClearParameter
                    
                    
        '\\Valida se existe cliente cadastrado com o CPF
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE socialSecurity = @socialSecutiry"
        .AddParameter "@socialSecutiry", TextSocialSecurity.Text, adVarChar
        
        Call .ExecuteWithQuery(Query.ToString())
        
        If Not .RecordSet.EOF Then
            If TextSocialSecurity.Text <> Empty And .RecordSet.Fields("Id").Value > 0 And .RecordSet.Fields("Id") <> Id Then
                  MsgBox "Um cliente cadastrado com esse CPF já existe!", _
                  vbExclamation, "JÁ EXISTE"
                  TextSocialSecurity.SetFocus
                  ValidateToUpdate = False
                Exit Function
            End If
        End If
        Query.Clear: .ClearParameter
        
        
        '\\Valida se existe cliente cadastrado com o RG
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE indentyCard = @indentyCard"
        .AddParameter "@indentyCard", TextIndentyCard.Text, adVarChar
        
        Call .ExecuteWithQuery(Query.ToString())
        
        If Not .RecordSet.EOF Then
            If TextIndentyCard.Text <> Empty And .RecordSet.Fields("Id").Value > 0 And .RecordSet.Fields("Id") <> Id Then
                  MsgBox "Um cliente cadastrado com esse RG já existe!", _
                  vbExclamation, "JÁ EXISTE"
                  TextIndentyCard.SetFocus
                  ValidateToUpdate = False
                Exit Function
            End If
        End If
        Query.Clear: .ClearParameter
        
        
        '\\Valida se existe cliente cadastrado com o email
        Query.Append "SELECT id FROM ClientNaturalPerson "
        Query.Append "WHERE email = @email"
        .AddParameter "@email", TextEmail.Text, adVarChar
        
        Call .ExecuteWithQuery(Query.ToString())
        
        If Not .RecordSet.EOF Then
            If TextEmail.Text <> Empty And .RecordSet.Fields("Id").Value > 0 And .RecordSet.Fields("Id") <> Id Then
                  MsgBox "Um cliente cadastrado com esse email já existe!", _
                  vbExclamation, "JÁ EXISTE"
                  TextEmail.SetFocus
                  ValidateToUpdate = False
                Exit Function
            End If
        End If
        Query.Clear: .ClearParameter
             
        ValidateToUpdate = True
        
    End With
    
    Set Query = Nothing
    
End Function
Private Function Update(ParameterId As Long) As Boolean

    If ValidateToUpdate() = False Then Exit Function

      Dim Query As StringBuilder
      Set Query = New StringBuilder
      
      With Query
        .Append "UPDATE ClientNaturalPerson Set internalCode = @internalCode, photoNumber = @photoNumber, activeStatus = @activeStatus,"
        .Append "registrationDate = @registrationDate, observation  = @observation, yourName = @yourName, age = @age, sex = @sex,"
        .Append "birthDay = @birthDay, civilStatus = @civilStatus, socialSecurity = @socialSecurity, indentyCard = @indentyCard,"
        .Append "fixedPhone = @fixedPhone, mobilePhone = @mobilePhone, whatsApp = @whatsApp, email = @email, district = @district,"
        .Append "city = @city, state = @state, zipCode = @zipCode, streetNumber = @streetNumber, addressDescription = @addressDescription,"
        .Append "addressComplement = @addressComplement WHERE id = @id"
      End With
         
      With New ConnectionAccess
      
        '\\Cliente
          .AddParameter "@internalCode", TextInternalCode.Text, adVarChar
          .AddParameter "@photoNumber", PhotoNumber, adVarChar
          .AddParameter "@activeStatus", TextActiveStatus.Text, adVarChar
          .AddParameter "@registrationDate", Date, adDate
          .AddParameter "@observation", TextObservation.Text, adVarChar
          .AddParameter "@yourName", TextYourName.Text, adVarChar
          .AddParameter "@age", TextAge.Text, adNumeric
          .AddParameter "@sex", TextSex.Text, adVarChar
          .AddParameter "@birthDay", TextBirthDay.Text, adDate
          .AddParameter "@civilStatus", TextCivilStatus.Text, adVarChar
          .AddParameter "@socialSecurity", TextSocialSecurity.Text, adVarChar
          .AddParameter "@indentyCard", TextIndentyCard.Text, adVarChar
                                       
        '\\Contato
          .AddParameter "@fixedPhone", TextFixedPhone.Text, adVarChar
          .AddParameter "@mobilePhone", TextMobilePhone.Text, adVarChar
          .AddParameter "@whatsApp", TextWhatsapp.Text, adVarChar
          .AddParameter "@email", TextEmail.Text, adVarChar
                                       
        '\\Endereço
          .AddParameter "district", TextDistrict.Text, adVarChar
          .AddParameter "city", TextCity.Text, adVarChar
          .AddParameter "state", TextState.Text, adVarChar
          .AddParameter "zipCode", TextZipCode.Text, adVarChar
          .AddParameter "streetNumber", TextStreetNumber.Text, adVarChar
          .AddParameter "addressDescription", TextAddressDescription.Text, adVarChar
          .AddParameter "addressComplement", TextAddressComplement.Text, adVarChar
          .AddParameter "@id", ParameterId, adNumeric
          
          Update = .ExecuteNonQuery(Query.ToString())
          
      End With
   
   Set Query = Nothing
   
End Function


Private Function Delete(ParameterId As Long) As Boolean

    Dim Query As StringBuilder
    
    If ParameterId = 0 Then
          MsgBox "Selecione um registro para excluir!", _
          vbExclamation, "SELECIONE"
        Exit Function
    End If
    
    If MsgBox("Confirma a exclusão desse cliente?", _
    vbExclamation + vbYesNo + vbDefaultButton2, _
    "IMPORTANTE") = vbNo Then Exit Function

    Set Query = New StringBuilder
        Query.Append "DELETE FROM ClientNaturalPerson WHERE id = @id"

    With New ConnectionAccess
        .AddParameter "@id", ParameterId, adNumeric
        Delete = .ExecuteNonQuery(Query.ToString())
    End With
     
    Set Query = Nothing
   
End Function


Public Sub ViewData(ParameterId As Long)

   On Error GoTo Exception
        
        Dim Query As StringBuilder
        
        If ParameterId = 0 Then
              MsgBox "Selecione um registro para visualizar!", vbExclamation, "SELECIONE"
           Exit Sub
        End If
           
        Set Query = New StringBuilder
            Query.Append "SELECT * FROM ClientNaturalPerson WHERE Id = @Id"
            
        With New ConnectionAccess
              
              '\\Adiciona parâmetro e executa query
              .AddParameter "@Id", ParameterId, adNumeric
              .ExecuteWithQuery Query.ToString()
          
                If .RecordSet.EOF = False Then
                
                      '\\Cliente
                      Id = .Field("id")
                      TextInternalCode.Text = .Field("internalCode")
                      PhotoNumber = .Field("photoNumber")
                      RegistrationDate = .Field("registrationDate")
                      TextActiveStatus.Text = .Field("ActiveStatus")
                      TextObservation.Text = .Field("observation")
                      TextYourName.Text = .Field("yourName")
                      TextAge.Text = .Field("age")
                      TextSex.Text = .Field("sex")
                      TextBirthDay.Text = .Field("birthDay")
                      TextCivilStatus.Text = .Field("civilStatus")
                      TextIndentyCard.Text = .Field("indentyCard")
                      TextSocialSecurity.Text = .Field("socialSecurity")
                      
                      '\\Contato
                      TextFixedPhone.Text = .Field("fixedPhone")
                      TextMobilePhone.Text = .Field("mobilePhone")
                      TextWhatsapp.Text = .Field("whatsApp")
                      TextEmail = .Field("email")
                      
                      '\\Endereço
                      TextDistrict.Text = .Field("district")
                      TextCity.Text = .Field("city")
                      TextState.Text = .Field("state")
                      TextZipCode.Text = .Field("zipCode")
                      TextStreetNumber.Text = .Field("streetNumber")
                      TextAddressDescription.Text = .Field("addressDescription")
                      TextAddressComplement.Text = .Field("addressComplement")
                      
                      With New Photograph
                         Call .LoadFile(ImageClient, CurrentPhoto)
                      End With
                      
                      Me.Show
                      
                End If
                
        End With
        
        Set Query = Nothing
                        
      Exit Sub

Exception:

   Call SysMethods.SubmitException

End Sub

Private Property Get CurrentPhoto() As String
   CurrentPhoto = SysDirectorys.PathUserFileClientPhoto & "\" & PhotoNumber & ".jpg"
End Property

Private Property Get ImageNothing() As String
   ImageNothing = SysDirectorys.PathAppFileIcon & "\ImageNothing.jpg"
End Property



