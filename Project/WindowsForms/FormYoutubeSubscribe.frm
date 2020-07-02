VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormYoutubeSubscribe 
   Caption         =   "Diário Excel"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "FormYoutubeSubscribe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormYoutubeSubscribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonSubscribe_Click()
   On Error GoTo error
   
         Dim TextFile As Variant
         
         With New FileSystemObject
            Set TextFile = .OpenTextFile(SysDirectorys.PathAppDef & "\Def.txt", ForWriting)
            
            With TextFile
               .WriteLine "YoutubeSubscrib = Ok"
               .WriteLine "App = Ok"
               .Close
            End With

         End With
         
         Call SysMethods.OpenPageWeb(SysPropertys.YoutubeChannel)

      Exit Sub
error:
      Unload Me
      SysMethods.SubmitException
End Sub

