Attribute VB_Name = "SysCollections"
Option Explicit

Public Sub SetYesNo(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Sim")
      .AddItem ("Não")
      .Text = .List(0)
   End With
End Sub

Public Sub SetSexes(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Masculino")
      .AddItem ("Feminino")
      .Text = .List(0)
   End With
End Sub

Public Sub SetCivilStatus(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Casado(a)")
      .AddItem ("Solteiro(a)")
      .AddItem ("Divorciado(a)")
      .Text = .List(0)
   End With
End Sub

Public Sub SetStatesLocation(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
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
      .Text = .List(0)
   End With
End Sub

Public Sub SetClientTypes(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Física")
       .AddItem ("Jurídica")
       .Text = .List(0)
   End With
End Sub

Public Sub SetCompanyTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Matriz")
       .AddItem ("Filial")
       .Text = .List(0)
   End With
End Sub

Public Sub SetCompanyTypeActions(ByVal ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Empresário Individual")
      .AddItem ("Microempreendedor - MEI")
      .AddItem ("Empresa Individual - EIRELI")
      .AddItem ("Sociedade Empresária")
      .AddItem ("Sociedade Simples")
      .Text = .List(0)
   End With
End Sub

