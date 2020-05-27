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
      .AddItem ("Amap�")
      .AddItem ("Amazonas")
      .AddItem ("Bahia")
      .AddItem ("Cear�")
      .AddItem ("Distrito Federal")
      .AddItem ("Espirito Santo")
      .AddItem ("Goi�s")
      .AddItem ("Maranh�o")
      .AddItem ("Mato Grosso")
      .AddItem ("Mato Grosso Do Sul")
      .AddItem ("Minas Gerais")
      .AddItem ("Par�")
      .AddItem ("Para�ba")
      .AddItem ("Paran�")
      .AddItem ("Pernanbuco")
      .AddItem ("Piau�")
      .AddItem ("Rio Grande Do Norte")
      .AddItem ("Rio Grande Do Sul")
      .AddItem ("Rio De Janeiro")
      .AddItem ("Rond�nia")
      .AddItem ("Roraima")
      .AddItem ("Santa Catarina")
      .AddItem ("S�o Paulo")
      .AddItem ("Sergipe")
      .AddItem ("Tocantins")
      .Text = "Select"
   End With
End Function

Public Function ListCustomerTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Select")
       .AddItem ("F�sica")
       .AddItem ("Jur�dica")
       .Text = "Select"
   End With
End Function

Public Function ListCompanyTypes(ComboBox As MSForms.ComboBox)
   With ComboBox
       .AddItem ("Select")
       .AddItem ("Matriz")
       .AddItem ("Filial")
       .AddItem ("�nico")
   End With
End Function

Public Function ListCompanyTypeActions(ComboBox As MSForms.ComboBox)
   With ComboBox
      .AddItem ("Select")
      .AddItem ("Empres�rio Individual")
      .AddItem ("Microempreendedor - MEI")
      .AddItem ("Empresa Individual - EIRELI")
      .AddItem ("Sociedade Empres�ria")
      .AddItem ("Sociedade Simples")
      .Text = "Select"
   End With
End Function