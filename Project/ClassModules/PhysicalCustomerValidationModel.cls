VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhysicalCustomerValidationModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ValidateToSave(Customer As PhysicalCustomerModel) As Boolean
   If Len(Customer.InternalCode) < 8 Then
      MsgBox "O código do cliente, deve ter no minimo 8 digitos!", vbExclamation, "Obrigatório"
      Exit Function
   End If
   
   If Customer.YourName = Empty Then
      MsgBox "O nome do cliente não foi informado!", vbExclamation, "Obrigatório"
      Exit Function
   End If
   ValidateToSave = True
End Function

Public Function ValidateToGetDetails(Id As Integer) As Boolean
   If Id = 0 Then
      MsgBox "Selecione um registro para visualizar!", vbExclamation, "Selecione"
      Exit Function
   End If
   ValidateToGetDetails = True
End Function

Public Function ValidateToDeleteSeveral(IdList As List) As Boolean
   If IdList.Count = 0 Then
      MsgBox "Nenhum registro da lista foi selecionado!", vbExclamation, "Selecione"
      Exit Function
   End If
   ValidateToDeleteSeveral = True
End Function

Public Function ValidateToDeleteUnic(Id As Integer) As Boolean
   If Id = 0 Then
      MsgBox "Selecione o registro para deletar!", vbExclamation, "Selecione"
      Exit Function
   End If
   ValidateToDeleteUnic = True
End Function
