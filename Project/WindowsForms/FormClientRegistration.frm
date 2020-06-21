VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormClientRegistration 
   Caption         =   "Diário Excel - Sistema Para Cadastro De Clientes"
   ClientHeight    =   10860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15975
   OleObjectBlob   =   "FormClientRegistration.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FormClientRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Class
   Id As Integer
   FileString As String
   PhotoNumber As String
End Type

Private This As Class
Private Mask() As New FormatterMask

Private Sub UserForm_Initialize()
   
   Call FormatMask
   Call FillComboBoxes
   Call SysFunction.DefineUserFormStyle(Me)
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
               This.PhotoNumber = Number.Generate(25, True)
            End If
            
            This.FileString = Picture.FileString
            
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

   Call SysFunction.SubmitException
   
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

   Call SysFunction.SubmitException
   
End Sub

Private Sub ButtonSave_Click()

   On Error GoTo Exception
      
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
      
      Dim Client As PhysicalClient
      Dim Picture As Photograph
      
      Set Client = New PhysicalClient
      Set Picture = New Photograph
      
      With Client
         .Id = This.Id
         .InternalCode = TextInternalCode.Value
         .YourName = TextYourName.Value
         .Age = TextAge.Value
         .BirthDay = TextBirthDay.Value
         .Sex = BoxSexes.Value
         .IndentyCard = TextIndentyCard.Value
         .SocialSecurity = TextSocialSecurity.Value
         .CivilStatus = BoxCivilStatus.Value
         .PhotoNumber = This.PhotoNumber
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
         .ActiveStatus = BoxActiveStatus.Value
      End With
         
      Select Case This.Id
         Case Is = 0
            If Client.Insert = True Then
               Call Picture.CopyFile(This.FileString, CurrentPhoto)
               Call ButtonClear_Click
               Call MsgBox("Registrado com sucesso!", vbInformation, "SUCESSO")
            End If
         Case Is > 0
            If Client.Update() = True Then
               Call Picture.CopyFile(This.FileString, CurrentPhoto)
               Call MsgBox("Editado com sucesso!", vbInformation, "SUCESSO")
            End If
      End Select

      Set Client = Nothing
      Set Picture = Nothing

      Exit Sub
      
Exception:

   Call SysFunction.SubmitException
    
End Sub

Private Sub ButtonClear_Click()

   On Error GoTo Exception
  
      Dim Control As Control
      
      For Each Control In Me.Controls
         
         Select Case TypeName(Control)
            
            Case Is = "TextBox"
               Control.Text = Empty
            
            Case Is = "ComboBox"
               Control.Text = "Selecionar"
            
            Case Is = "CheckBox"
               Control.Value = False
                        
            Case Is = "Image"
               Call LoadImageNothing(Control)
               
         End Select
      Next
      
      With TextInternalCode
         .Locked = False
      End With
      
      With BoxActiveStatus
         .SetFocus
      End With
      
      With This
         .Id = 0
         .FileString = Empty
         .PhotoNumber = Empty
      End With
   
   Exit Sub

Exception:

   Call SysFunction.SubmitException
   
End Sub

Private Sub ButtonDelete_Click()

   On Error GoTo Exception
      
      If This.Id = 0 Then
            MsgBox "Selecione o registro para deletar!", vbExclamation, "Selecione"
         Exit Sub
      End If
      
      Dim Client As PhysicalClient
      Dim Picture As Photograph
         
      Set Client = New PhysicalClient
      Set Picture = New Photograph
      
      If Client.Delete(This.Id) = True Then
         
         Call Picture.DeleteFile(CurrentPhoto)
         Call ButtonClear_Click
            
         MsgBox "Deletado com sucesso!", vbInformation, "Sucesso"
      
      End If
      
      Set Client = Nothing
      Set Picture = Nothing
      
   Exit Sub
   
Exception:

   Call SysFunction.SubmitException
   
End Sub

Private Sub ButtonClose_Click()
   Unload Me
End Sub

Private Sub LoadImageNothing(Image As MSForms.Image)
   
   On Error GoTo Exception
   
      With New Photograph
         Call .LoadFile(Image, ImageNothing)
      End With
      
      This.FileString = Empty
      This.PhotoNumber = Empty
   
   Exit Sub

Exception:
      
   Call SysFunction.SubmitException
   
End Sub

Private Sub FillComboBoxes()

   On Error GoTo Exception
   
      With New CollectionTypes
          Call .ListSexes(BoxSexes)
          Call .ListCivilStatus(BoxCivilStatus)
          Call .ListStates(BoxStates)
          Call .ListYesNo(BoxActiveStatus)
      End With
      
   Exit Sub
   
Exception:

   Call SysFunction.SubmitException
   
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
         End Select
      Next
      
   Exit Sub
   
Exception:

   Call SysFunction.SubmitException
   
End Sub

Public Sub SetDetails(Id As Integer)

   On Error GoTo Exception
        
      If Id = 0 Then
            MsgBox "Selecione um registro para visualizar!", vbExclamation, "Selecione"
         Exit Sub
      End If
        
      Dim Client As PhysicalClient
      Dim Image As MSForms.Image
               
      Set Client = New PhysicalClient
      Set Image = ImageClient
      
         Call Client.GetDetails(Id)
         
         With Client
            This.Id = .Id
            TextInternalCode.Value = .InternalCode
            TextYourName.Value = .YourName
            TextAge.Value = .Age
            TextBirthDay.Value = .BirthDay
            BoxSexes.Value = .Sex
            TextIndentyCard.Value = .IndentyCard
            TextSocialSecurity.Value = .SocialSecurity
            BoxCivilStatus.Value = .CivilStatus
            This.PhotoNumber = .PhotoNumber
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
            BoxActiveStatus.Value = .ActiveStatus
         End With
         
         With New Photograph
            Call .LoadFile(Image, CurrentPhoto)
         End With
   
      Set Client = Nothing
      
      Me.Show
      
      Exit Sub
      
Exception:

   Call SysFunction.SubmitException
      
End Sub

Private Property Get CurrentPhoto() As String
   CurrentPhoto = SysProperty.ClientPhotosDirectory & "\" & This.PhotoNumber & ".jpg"
End Property

Private Property Get ImageNothing() As String
   ImageNothing = SysProperty.AppFileIconsDirectory & "\ImageNothing.jpg"
End Property
