VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExceptionErrorNotifier 
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   OleObjectBlob   =   "FormExceptionErrorNotifier.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExceptionErrorNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
                   
      With New StringBuilder
         
         .Append "Ocorreu uma falha durante o processamento dessa operação, "
         .Append "e por esse motivo foi gerada uma exceção."
         
         .Append vbCr
         .Append vbCr
         
         .Append "-----------------------------------------------" & vbCr
         .Append "Informações sobre o erro" & vbNewLine
         .Append "-----------------------------------------------" & vbCr
         
         .Append vbCr
         .Append Err.Description
         
         .Append vbCr
         .Append vbCr
         
         .Append "Número: " & Err.Number
         
         .Append vbCr
         .Append vbCr
         .Append "-----------------------------------------------" & vbCr
         
         .Append vbCr
         .Append "Acesse: diarioexcel.com.br e solicite nosso suporte!"
         
          Me.TextErrorDescription.Text = .ToString()
          
      End With
      
      
     '\\Salva o log do erro
      Dim ErrLog As String
      ErrLog = "Erro:" & Err.Number & Space(1) & "Description:" & Err.Description & _
      Space(1) & "Data:" & Now & vbCrLf

      Call SaveLog(SysDirectorys.PathAppLog & "\Error.log", ErrLog)
    
End Sub
