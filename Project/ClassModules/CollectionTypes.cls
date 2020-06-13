VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListYesNo(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
      .AddItem ("Sim")
      .AddItem ("Não")
      .Text = "Select"
   End With
End Function

Public Function ListSexes(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
      .AddItem ("Masculino")
      .AddItem ("Feminino")
      .Text = "Select"
   End With
End Function

Public Function ListCivilStatus(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
      .AddItem ("Casado(a)")
      .AddItem ("Solteiro(a)")
      .AddItem ("Divorciado(a)")
      .Text = "Select"
   End With
End Function

Public Function ListStates(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
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
      .Text = "Select"
   End With
End Function

Public Function ListCustomerTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Select")
       .AddItem ("Física")
       .AddItem ("Jurídica")
       .Text = "Select"
   End With
End Function

Public Function ListCompanyTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Select")
       .AddItem ("Matriz")
       .AddItem ("Filial")
       .AddItem ("Único")
   End With
End Function

Public Function ListCompanyTypeActions(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
      .AddItem ("Empresário Individual")
      .AddItem ("Microempreendedor - MEI")
      .AddItem ("Empresa Individual - EIRELI")
      .AddItem ("Sociedade Empresária")
      .AddItem ("Sociedade Simples")
      .Text = "Select"
   End With
End Function
