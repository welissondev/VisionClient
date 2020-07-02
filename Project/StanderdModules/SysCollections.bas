Attribute VB_Name = "SysCollections"
Option Explicit

Public Sub SetYesNo(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Selecionar")
      .AddItem ("Sim")
      .AddItem ("Não")
      .Text = "Selecionar"
   End With
End Sub

Public Sub SetSexes(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Selecionar")
      .AddItem ("Masculino")
      .AddItem ("Feminino")
      .Text = "Selecionar"
   End With
End Sub

Public Sub SetCivilStatus(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Selecionar")
      .AddItem ("Casado(a)")
      .AddItem ("Solteiro(a)")
      .AddItem ("Divorciado(a)")
      .Text = "Selecionar"
   End With
End Sub

Public Sub SetStatesLocation(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Selecionar")
      .AddItem ("Acre")
      .AddItem ("Alagoas")
      .AddItem ("Amapá")
      .AddItem ("Amazonas")
      .AddItem ("Bahia")
      .AddItem ("Ceará")
      .AddItem ("Distrito Federal")
      .AddItem ("Espirito Santo")
      .AddItem ("Goiás")
      .AddItem ("Maranhão")
      .AddItem ("Mato Grosso")
      .AddItem ("Mato Grosso Do Sul")
      .AddItem ("Minas Gerais")
      .AddItem ("Pará")
      .AddItem ("Paraíba")
      .AddItem ("Paraná")
      .AddItem ("Pernanbuco")
      .AddItem ("Piauí")
      .AddItem ("Rio Grande Do Norte")
      .AddItem ("Rio Grande Do Sul")
      .AddItem ("Rio De Janeiro")
      .AddItem ("Rondônia")
      .AddItem ("Roraima")
      .AddItem ("Santa Catarina")
      .AddItem ("São Paulo")
      .AddItem ("Sergipe")
      .AddItem ("Tocantins")
      .Text = "Selecionar"
   End With
End Sub

Public Sub SetClientTypes(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Selecionar")
       .AddItem ("Física")
       .AddItem ("Jurídica")
       .Text = "Selecionar"
   End With
End Sub

Public Sub SetCompanyTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Selecionar")
       .AddItem ("Matriz")
       .AddItem ("Filial")
       .AddItem ("Único")
   End With
End Sub

Public Sub SetCompanyTypeActions(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Selecionar")
      .AddItem ("Empresário Individual")
      .AddItem ("Microempreendedor - MEI")
      .AddItem ("Empresa Individual - EIRELI")
      .AddItem ("Sociedade Empresária")
      .AddItem ("Sociedade Simples")
      .Text = "Selecionar"
   End With
End Sub

