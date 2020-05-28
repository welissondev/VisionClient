Attribute VB_Name = "AppConfig"
Option Explicit

Public Property Get ConnectionStrig() As String
   ConnectionStrig = "Provider=Microsoft.ACE.Oledb.12.0;Data Source = " & ThisWorkbook.Path & "\App\Data\VisionBase.mdb"
End Property

Public Property Get ClientPhotosDirectory() As String
   ClientPhotosDirectory = ThisWorkbook.Path & "\User\Vision\ClientPhotos"
End Property

Public Property Get AppFileIconsDirectory() As String
   AppFileIconsDirectory = ThisWorkbook.Path & "\App\File\Icons"
End Property

Public Property Get AppVersion() As String
   AppVersion = "v2.0.13.100"
End Property

Public Property Get AppAName() As String
   AppAName = "Vision Client"
End Property

Public Property Get CompanyName() As String
   CompanyName = "Diário Excel"
End Property

Public Sub SetScreenControlStyle(Form As MSForms.UserForm)
   Dim x As Control
   For Each x In Form.Controls
      
      Form.BackColor = &H996600
      
      Select Case TypeName(x)
         Case Is = "TextBox"
            x.BorderColor = 14540253
            x.BackColor = 15395562
            x.ForeColor = 1842204
            x.BorderStyle = 1
            x.FontName = "Comic sans ms"
         Case Is = "ComboBox"
            x.BorderColor = 14540253
            x.BackColor = 15395562
            x.ForeColor = 1842204
            x.BorderStyle = 1
            x.FontName = "Comic sans ms"
         Case Is = "Label"
            x.ForeColor = Form.BackColor
            x.FontName = "Comic sans ms"
         Case Is = "Frame"
            x.BorderColor = 14540253
            x.BackColor = &HFFFFFF
         Case Is = "OptionButton"
            x.ForeColor = Form.BackColor
            x.FontName = "Comic sans ms"
      End Select
   Next
End Sub
