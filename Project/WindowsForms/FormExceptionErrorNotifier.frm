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
Option Explicit

Private Sub UserForm_Initialize()
  SetError
End Sub

Private Sub SetError()
   
   Dim ErrorStrBilder As StringBuilder
   Dim Sheet As Worksheet
   Dim Row As Long
   
   Set ErrorStrBilder = New StringBuilder
      With ErrorStrBilder
         
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
         
      End With
      
      ErrorCaption = "Erro Em Tempo De Execução :("
      ErrorDescription = ErrorStrBilder.ToString
      
      Err.Clear
      
End Sub

