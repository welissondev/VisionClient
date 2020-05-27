VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhysicalCustomerAddScreen 
   Caption         =   "Diário Excel - Sistema Para Cadastro De Clientes"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   OleObjectBlob   =   "PhysicalCustomerAddScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "PhysicalCustomerAddScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Class
   Id As Integer
End Type

Private This As Class

Private PhotoString As String
Private PhotoNumber As String
Private ActiveStatus As Boolean
Private Mask() As New FormatterMask

Private Sub ButtonSelectPhoto_Click()
   On Error GoTo Error
      With New PictureFile
         If .CheckFile(.OpenFile) = True Then
            .LoadFile ImageCustomer
             PhotoString = .FileString
            If CheckIfPhotoExists(PhotoNumber) = False Then
               PhotoNumber = .FileNumber
            End If
         Else
            If CheckIfPhotoExists(PhotoNumber) = False Then
               .LoadFile ImageCustomer, ImageNothing
               PhotoString = vbNullString
               PhotoNumber = vbNullString
            End If
         End If
      End With
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Function CheckIfPhotoExists(NumberId As String) As Boolean
   With New PictureFile
      CheckIfPhotoExists = .CheckFile(DirectoryPhoto & NumberId & ".jpg")
   End With
End Function

Private Sub SavePhoto()
   On Error GoTo Error
      With New PictureFile
         Call .CopyFile(PhotoString, DirectoryPhoto & PhotoNumber & ".jpg")
      End With
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Sub UserForm_Initialize()
   Call FormatMask
   Call FillAllComboBox
End Sub

Private Sub CheckGenerateCode_Click()
   Select Case CheckGenerateCode.Value
      Case Is = True
         TextInternalCode.Text = SingleCode.GetCode(8)
         TextInternalCode.Locked = True
         TextYourName.SetFocus
      Case Is = False
         TextInternalCode.Text = Empty
         TextInternalCode.Locked = False
         TextInternalCode.SetFocus
   End Select
End Sub

Private Sub ButtonSave_Click()
   On Error GoTo Error
   
      Dim Customer As PhysicalCustomerModel
      
      Set Customer = New PhysicalCustomerModel
         With Customer
            .InternalCode = TextInternalCode.Value
            .YourName = TextYourName.Value
            .Age = TextAge.Value
            .BirthDay = TextBirthDay.Value
            .Sex = BoxSexes.Value
            .IndentyCard = TextIndentyCard.Value
            .SocialSecurity = TextSocialSecurity.Value
            .CivilStatus = BoxCivilStatus.Value
            .PhotoNumber = PhotoNumber
            .FixedPhone = TextFixedPhone.Value
            .MobilePhone = TextMobilePhone.Value
            .WhatsApp = TextWhatsapp.Value
            .Email = TextEmail.Value
            .AddressDescription = TextAddressDescription.Value
            .AddressComplement = TextAddressComplement.Value
            .AddressNote = TextAddressNote.Value
            .District = TextDistrict.Value
            .City = TextCity.Value
            .State = BoxStates.Value
            .ZipCode = TextZipCode.Value
            .StreetNumber = TextStreetNumber.Value
            .ActiveStatus = ActiveStatus
         End With
            
         Select Case This.Id
            Case Is = 0
               If Customer.Insert = True Then
                  Call SavePhoto
                  Call ResetScreen
                  Call MsgBox("Registrado com sucesso!", vbInformation, "SUCESSO")
               End If
            Case Is > 0
               If Customer.Update(This.Id) = True Then
                  Call SavePhoto
                  Call MsgBox("Editado com sucesso!", vbInformation, "SUCESSO")
               End If
         End Select
   
         Set Customer = Nothing

      Exit Sub
Error:

    ErrorNoteScreen.Show
End Sub

Private Sub ButtonClear_Click()
   Call ResetScreen
End Sub

Private Sub ButtonDelete_Click()
   On Error GoTo Error
      
      Dim Customer As PhysicalCustomerModel
      Dim IDC As Collection
      
      Set IDC = New Collection
         With IDC
            .Add This.Id
         End With
         
      Set Customer = New PhysicalCustomerModel
         If Customer.Delete(IDC) = True Then
            With New PictureFile
               If .DeleteFile(DirectoryPhoto & PhotoNumber & ".jpg") = True Then
                  Call ResetScreen
                  MsgBox "Deletado com sucesso!", vbInformation, "Sucesso"
               End If
            End With
         End If
     Set Customer = Nothing
     
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Sub ButtonClose_Click()
   Unload Me
End Sub

Public Sub SetDetails(Id As Integer)
   On Error GoTo Error
        
      Dim Customer As PhysicalCustomerModel
            
      Set Customer = New PhysicalCustomerModel
         With Customer
            Call .GetDetails(Id)
         End With
      
         With Customer
            This.Id = .Id
            TextInternalCode.Value = .InternalCode
            TextYourName.Value = .YourName
            TextAge.Value = .Age
            TextBirthDay.Value = .BirthDay
            BoxSexes.Value = .Sex
            TextIndentyCard.Value = .IndentyCard
            TextSocialSecurity.Value = .SocialSecurity
            BoxCivilStatus.Value = .CivilStatus
            PhotoNumber = .PhotoNumber
            TextFixedPhone.Value = .FixedPhone
            TextMobilePhone.Value = .MobilePhone
            TextWhatsapp.Value = .WhatsApp
            TextEmail.Value = .Email
            TextAddressDescription.Value = .AddressDescription
            TextAddressComplement.Value = .AddressComplement
            TextAddressNote.Value = .AddressNote
            TextDistrict.Value = .District
            TextCity.Value = .City
            BoxStates.Value = .State
            TextZipCode.Value = .ZipCode
            TextStreetNumber.Value = .StreetNumber
            ActiveStatus = .ActiveStatus
         End With
         
         PhotoString = DirectoryPhoto & PhotoNumber & ".jpg"
         
         With New PictureFile
            .FileString = PhotoString
            Call .LoadFile(ImageCustomer)
         End With
         
         Select Case ActiveStatus
            Case Is = True
               OptionInactive.Value = False
               OptionActive.Value = True
            Case Is = False
               OptionActive.Value = False
               OptionInactive.Value = True
         End Select
   
      Set Customer = Nothing
      
      Exit Sub
Error:
      ErrorNoteScreen.Show
End Sub

Private Sub ResetScreen()
   On Error GoTo Error
     
      Dim x As Control
      For Each x In Me.Controls
         Select Case TypeName(x)
            Case Is = "TextBox"
               x.Text = Empty
            
            Case Is = "ComboBox"
               x.Text = "Select"
            
            Case Is = "CheckBox"
               x.Value = False
            
            Case Is = "OptionButton"
               x.Value = False
               If x.Name = "OptionInactive" Then
                  x.Value = True
               End If
            
            Case Is = "Image"
               x.Picture = LoadPicture(ImageNothing)
         End Select
      Next
      
      With TextInternalCode
         .Locked = False
         .SetFocus
      End With
      
      This.Id = 0
      PhotoString = Empty
      PhotoNumber = Empty
      ActiveStatus = False
      
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Sub FillAllComboBox()
   On Error GoTo Error
      With New CollectionTypes
          Call .ListSexes(BoxSexes)
          Call .ListCivilStatus(BoxCivilStatus)
          Call .ListStates(BoxStates)
      End With
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Sub FormatMask()
   On Error GoTo Error
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
            Case Is = "Number"
               Set Mask(Index).CanNotString = Controls(Index)
            Case Is = "ZipCode"
               Set Mask(Index).ToZipCode = Controls(Index)
         End Select
      Next
   Exit Sub
Error:
   ErrorNoteScreen.Show
End Sub

Private Sub OptionActive_Click()
   OptionInactive = False: OptionActive = True: ActiveStatus = True
End Sub
Private Sub OptionInactive_Click()
   OptionInactive = True: OptionActive = False: ActiveStatus = False
End Sub


Private Property Get ImageNothing() As String
   ImageNothing = ThisWorkbook.Path & "\App\File\Icons\ImageNothing.jpg"
End Property
Private Property Get DirectoryPhoto() As String
   DirectoryPhoto = ThisWorkbook.Path & "\User\Vision\ClientPhotos\"
End Property
