Attribute VB_Name = "AppFunction"
Option Explicit

'***************************************************
'Nesse módulo fica todos as funções que estão
'disponiveis publicamente para todo o sistema
'***************************************************

Public Sub SetStyle(Form As MSForms.UserForm)
   
   Dim Control As MSForms.Control

   Dim TextBorderColor, TextBackColor, TextForeColor, TextBorderStyle, TextFont As Variant
      TextFont = "Arial"
      TextBackColor = 15395562
      TextBorderColor = 14540253
      TextForeColor = 1842204
      TextBorderStyle = 1
   
   Dim LabelForeColor, LabelFont As Variant
      LabelForeColor = &H996600
      LabelFont = "Arial"
      
   Dim FrameBackColor, FrameForeColor, FrameBorderColor, FrameFont As Variant
      FrameBorderColor = 14540253
      FrameForeColor = 1842204
      FrameBackColor = &HFFFFFF
      FrameFont = "Arial"
   
   Form.Font = "Arial"
   Form.BackColor = &H996600
   
   For Each Control In Form.Controls
      
      Select Case TypeName(Control)
         Case Is = "TextBox"
            With Control
               .BorderColor = TextBorderColor
               .BackColor = TextBackColor
               .ForeColor = TextForeColor
               .BorderStyle = TextBorderStyle
               With .Font
                  .Name = TextFont
                  .Size = 10
               End With
            End With
         Case Is = "ComboBox"
            With Control
               .BorderColor = TextBorderColor
               .BackColor = TextBackColor
               .ForeColor = TextForeColor
               .BorderStyle = TextBorderStyle
               With .Font
                  .Name = TextFont
                  .Size = 10
               End With
            End With
            
         Case Is = "Label"
            With Control
               .ForeColor = LabelForeColor
               With .Font
                  .Name = LabelFont
                  .Size = 10
               End With
            End With
            
         Case Is = "Frame"
            With Control
               .BorderColor = FrameBorderColor
               .BackColor = FrameBackColor
               .ForeColor = FrameForeColor
               With .Font
                  .Name = FrameFont
                  .Size = 10
               End With
            End With
            
         Case Is = "OptionButton"
            With Control
               .ForeColor = Form.BackColor
               With .Font
                  .Name = TextFont
                  .Size = 10
               End With
            End With
      End Select
   Next
   
End Sub


