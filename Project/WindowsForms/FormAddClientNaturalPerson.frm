VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAddClientNaturalPerson 
   Caption         =   "Diário Excel - Sistema Para Cadastro De Clientes"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14895
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

Implements Client: Implements Person: Implements Contact: Implements Address

Private Id As Long
Private FileString As String
Private PhotoNumber As String
Private RegistrationDate As Date
Private Mask() As New FormatterMask


Private Sub UserForm_Initialize()

   Call FormatMask
   Call FillTextBox
   Call SysMethods.DefineUserFormStyle(Me, &H404040)
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
        
      Dim Client As ClientNaturalPerson
      Dim Picture As Photograph
        
      If TextActiveStatus.Text = "Selecionar" Then
            MsgBox "Defina o estado de atividade do cliente!", vbExclamation, "Obrigatório"
            TextActiveStatus.SetFocus
         Exit Sub
      End If
        
      If Len(TextInternalCode.Value) < 8 Then
            MsgBox "O código do cliente, deve ter no minimo 8 digitos!", vbExclamation, "Obrigatório"
            TextInternalCode.SetFocus
         Exit Sub
      End If

      If TextYourName.Value = Empty Then
            MsgBox "O nome do cliente não foi informado!", vbExclamation, "Obrigatório"
            TextYourName.SetFocus
         Exit Sub
      End If

      Set Client = New ClientNaturalPerson
      Set Picture = New Photograph
      
      Call Client.Builder(Me)
            
      Select Case Id
         Case Is = 0
            If Client.Insert = True Then
               Call Picture.CopyFile(FileString, CurrentPhoto)
               Call ButtonClear_Click
               Call MsgBox("Registrado com sucesso!", vbInformation, "SUCESSO")
            End If
         Case Is > 0
            If Client.Update() = True Then
               Call Picture.CopyFile(FileString, CurrentPhoto)
               Call MsgBox("Editado com sucesso!", vbInformation, "SUCESSO")
            End If
      End Select

      Set Client = Nothing
      Set Picture = Nothing

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
               Control.Value = "Selecionar"

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

      If Id = 0 Then
            MsgBox "Selecione o registro para deletar!", vbExclamation, "Selecione"
         Exit Sub
      End If

      Dim Client As ClientNaturalPerson
      Dim Picture As Photograph

      Set Client = New ClientNaturalPerson
      Set Picture = New Photograph

      If Client.Delete(Id) = True Then

         Call Picture.DeleteFile(CurrentPhoto)
         Call ButtonClear_Click

         MsgBox "Deletado com sucesso!", vbInformation, "Sucesso"

      End If

      Set Client = Nothing
      Set Picture = Nothing

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
               .Text = Code.Generate(8)
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
            Case Is = "InternalCode"
                Set Mask(Index).CanNotString = Controls(Index)
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



Public Sub ViewData(Id As Long)

   On Error GoTo Exception
        
        If Id = 0 Then
              MsgBox "Selecione um registro para visualizar!", vbExclamation, "Selecione"
           Exit Sub
        End If
  
        With New ClientNaturalPerson
           Call .Builder(Me): Call .ViewData(Id)
        End With
               
        Me.Show

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


'------------------------------------------------------------------
'Propriedades implementadas
'-----------------------------//-----------------------------------

'Entidade cliente
Private Property Get Client_Id() As Long
    Client_Id = Id
End Property
Private Property Let Client_Id(Value As Long)
    Id = Value
End Property
Private Property Get Client_InternalCode() As String
    Client_InternalCode = TextInternalCode.Text
End Property
Private Property Let Client_InternalCode(Value As String)
    TextInternalCode.Text = Value
End Property
Private Property Get Client_PhotoNumber() As String
    Client_PhotoNumber = PhotoNumber
End Property
Private Property Let Client_PhotoNumber(Value As String)
    PhotoNumber = Value
    With New Photograph
        Call .LoadFile(ImageClient, CurrentPhoto)
    End With
End Property
Private Property Get Client_RegistrationDate() As Date
    Client_RegistrationDate = RegistrationDate
End Property
Private Property Let Client_RegistrationDate(Value As Date)
    RegistrationDate = Value
End Property
Private Property Get Client_ActiveStatus() As String
    If TextActiveStatus.Text <> "Selecionar" Then
        Client_ActiveStatus = TextActiveStatus.Text
    End If
End Property
Private Property Let Client_ActiveStatus(Value As String)
    If Value <> Empty Then
        TextActiveStatus.Text = Value
    Else
        TextActiveStatus.Text = "Selecionar"
    End If
End Property
Private Property Get Client_Observation() As String
    Client_Observation = TextObservation.Text
End Property
Private Property Let Client_Observation(Value As String)
    TextObservation.Text = Value
End Property



'Entidade pessoa
Private Property Get Person_Age() As Integer
    If TextAge.Text <> Empty Then
        Person_Age = TextAge.Text
    End If
End Property
Private Property Let Person_Age(Value As Integer)
    If Value > 0 Then
        TextAge.Text = Value
    End If
End Property
Private Property Get Person_BirthDay() As Date
    If TextBirthDay.Text <> Empty Then
      Person_BirthDay = TextBirthDay.Text
    End If
End Property
Private Property Let Person_BirthDay(Value As Date)
    If Value <> "00:00:00" Then
        TextBirthDay = Value
    End If
End Property
Private Property Get Person_CivilStatus() As String
    If TextCivilStatus.Text <> "Selecionar" Then
        Person_CivilStatus = TextCivilStatus.Text
    End If
End Property
Private Property Let Person_CivilStatus(Value As String)
    If Value <> Empty Then
        TextCivilStatus.Text = Value
    Else
        TextCivilStatus.Text = "Selecionar"
    End If
End Property
Private Property Get Person_IndentyCard() As String
    Person_IndentyCard = TextIndentyCard.Text
End Property
Private Property Let Person_IndentyCard(Value As String)
    TextIndentyCard.Text = Value
End Property
Private Property Get Person_Name() As String
    Person_Name = TextYourName.Text
End Property
Private Property Let Person_Name(Value As String)
    TextYourName.Text = Value
End Property
Private Property Get Person_Sex() As String
    If TextSex.Text <> "Selecionar" Then
        Person_Sex = TextSex.Text
    End If
End Property
Private Property Let Person_Sex(Value As String)
    If Value <> Empty Then
        TextSex.Text = Value
    Else
        TextSex.Text = "Selecionar"
    End If
End Property
Private Property Get Person_SocialSecurity() As String
    Person_SocialSecurity = TextSocialSecurity.Text
End Property
Private Property Let Person_SocialSecurity(Value As String)
    TextSocialSecurity.Text = Value
End Property


'Entidade Contato
Private Property Get Contact_Email() As String
    Contact_Email = TextEmail.Text
End Property
Private Property Let Contact_Email(Value As String)
    TextEmail.Text = Value
End Property
Private Property Get Contact_FixedPhone() As String
    Contact_FixedPhone = TextFixedPhone.Text
End Property
Private Property Let Contact_FixedPhone(Value As String)
    TextFixedPhone.Text = Value
End Property
Private Property Get Contact_MobilePhone() As String
    Contact_MobilePhone = TextMobilePhone.Text
End Property
Private Property Let Contact_MobilePhone(Value As String)
    TextMobilePhone.Text = Value
End Property
Private Property Get Contact_WhatsApp() As String
    Contact_WhatsApp = TextWhatsapp.Text
End Property
Private Property Let Contact_WhatsApp(Value As String)
    TextWhatsapp.Text = Value
End Property


'Entidade endereço
Private Property Get Address_City() As String
    Address_City = TextCity.Text
End Property
Private Property Let Address_City(Value As String)
    TextCity = Value
End Property
Private Property Get Address_Complement() As String
    Address_Complement = TextAddressComplement.Text
End Property
Private Property Let Address_Complement(Value As String)
    TextAddressComplement.Text = Value
End Property
Private Property Get Address_Description() As String
    Address_Description = TextAddressDescription.Text
End Property
Private Property Let Address_Description(Value As String)
    TextAddressDescription.Text = Value
End Property
Private Property Get Address_District() As String
    Address_District = TextDistrict.Text
End Property
Private Property Let Address_District(Value As String)
    TextDistrict.Text = Value
End Property
Private Property Get Address_State() As String
    If TextState.Text <> "Selecionar" Then
        Address_State = TextState.Text
    End If
End Property
Private Property Let Address_State(Value As String)
    If Value <> Empty Then
        TextState.Text = Value
    Else
        TextState.Text = "Selecionar"
    End If
End Property
Private Property Get Address_StreetNumber() As String
    Address_StreetNumber = TextStreetNumber.Text
End Property
Private Property Let Address_StreetNumber(Value As String)
    TextStreetNumber.Text = Value
End Property
Private Property Get Address_ZipCode() As String
    Address_ZipCode = TextZipCode.Text
End Property
Private Property Let Address_ZipCode(Value As String)
    TextZipCode.Text = Value
End Property


