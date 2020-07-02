VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatterMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ToDate As MSForms.TextBox
Attribute ToDate.VB_VarHelpID = -1
Public WithEvents ToSocialSecurity As MSForms.TextBox
Attribute ToSocialSecurity.VB_VarHelpID = -1
Public WithEvents ToNationalRegistry As MSForms.TextBox
Attribute ToNationalRegistry.VB_VarHelpID = -1
Public WithEvents ToSocialSecutiryAndNationalRegistry As MSForms.TextBox
Attribute ToSocialSecutiryAndNationalRegistry.VB_VarHelpID = -1
Public WithEvents ToZipCode As MSForms.TextBox
Attribute ToZipCode.VB_VarHelpID = -1
Public WithEvents ToMobilePhone As MSForms.TextBox
Attribute ToMobilePhone.VB_VarHelpID = -1
Public WithEvents ToFixedPhone As MSForms.TextBox
Attribute ToFixedPhone.VB_VarHelpID = -1
Public WithEvents ToOnlyNull As MSForms.TextBox
Attribute ToOnlyNull.VB_VarHelpID = -1
Public WithEvents ToCurrency As MSForms.TextBox
Attribute ToCurrency.VB_VarHelpID = -1
Public WithEvents CanNotString  As MSForms.TextBox
Attribute CanNotString.VB_VarHelpID = -1
Public WithEvents ToJustDotString  As MSForms.TextBox
Attribute ToJustDotString.VB_VarHelpID = -1

Private Sub CanNotString_Change()
   Dim ret As String
   Dim I As Integer
       For I = 1 To Len(CanNotString.Value)
           If Asc(Mid(CanNotString.Value, I, 1)) >= 48 And Asc(Mid(CanNotString.Value, I, 1)) <= 58 Then
               ret = ret & Mid(CanNotString.Value, I, 1)
           End If
       Next I
   CanNotString = ret
End Sub
Private Sub ToJustDotString_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToJustDotString
       .MaxLength = 9
       Select Case KeyAscii
           Case 48 To 57
              If .SelStart = 6 Then .SelText = "."
           Case Else: KeyAscii = 0
       End Select
   End With
End Sub
Private Sub ToSocialSecurity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToSocialSecurity
      .MaxLength = 14
      Select Case KeyAscii
           Case 8
           Case 13: SendKeys "{TAB}"
           Case 48 To 57
              If .SelStart = 3 Then .SelText = "."
              If .SelStart = 7 Then .SelText = "."
              If .SelStart = 11 Then .SelText = "-"
           Case Else: KeyAscii = 0
        End Select
    End With
End Sub
Private Sub ToNationalRegistry_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToNationalRegistry
       .MaxLength = 18
        Select Case KeyAscii
             Case 8
             Case 13: SendKeys "{TAB}"
             Case 48 To 57
                If .SelStart = 2 Then .SelText = "."
                If .SelStart = 6 Then .SelText = "."
                If .SelStart = 10 Then .SelText = "/"
                If .SelStart = 15 Then .SelText = "-"
             Case Else: KeyAscii = 0
          End Select
     End With
End Sub
Private Sub ToSocialSecutiryAndNationalRegistry_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
       KeyAscii = 0
   End If
   With ToSocialSecutiryAndNationalRegistry
       If Len(.Text) = 2 Then
           .SelText = "."
       ElseIf Len(.Text) = 4 Then
           .Text = Replace(.Text, ".", "")
           .SelText = "."
       ElseIf Len(.Text) = 6 Then
           .Text = Replace(.Text, ".", "")
           .Text = Format(.Text, "00"".""000"".")
       ElseIf Len(.Text) = 8 Then
           .Text = Replace(.Text, ".", "")
           .Text = Format(.Text, "000"".""000"".")
       ElseIf Len(.Text) = 10 Then
           .Text = Replace(.Text, ".", "")
           .Text = Format(.Text, "00"".""000"".""000")
       ElseIf Len(.Text) = 11 Then
           .Text = Replace(.Text, ".", "")
           .Text = Format(.Text, "000"".""000"".""000""-")
       ElseIf Len(.Text) = 15 Then
           .Text = Replace(.Text, ".", "")
           .Text = Replace(.Text, "-", "")
           .Text = Format(.Text, "00"".""000"".""000""/""0000""-")
       End If
   End With
End Sub
Private Sub ToDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToDate
      .MaxLength = 10
      Select Case KeyAscii
         Case 8
         Case 13: SendKeys "{TAB}"
         Case 48 To 57
            If .SelStart = 2 Then .SelText = "/"
            If .SelStart = 5 Then .SelText = "/"
         Case Else: KeyAscii = 0
      End Select
   End With
End Sub
Private Sub ToZipCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToZipCode
       .MaxLength = 9
       Select Case KeyAscii
           Case 8
           Case 13: SendKeys "{TAB}"
           Case 48 To 57
              If .SelStart = 5 Then .SelText = "-"
           Case Else: KeyAscii = 0
       End Select
   End With
End Sub
Private Sub ToFixedPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   With ToFixedPhone
      .MaxLength = 14
      Select Case KeyAscii
           Case 8
           Case 13: SendKeys "{TAB}"
           Case 48 To 57
              If .SelStart = 0 Then .SelText = "("
              If .SelStart = 3 Then .SelText = ")" & Space(1)
              If .SelStart = 9 Then .SelText = "-"
           Case Else: KeyAscii = 0
        End Select
   End With
End Sub
Private Sub ToMobilePhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With ToMobilePhone
        .MaxLength = 15
        Select Case KeyAscii
             Case 8
             Case 13: SendKeys "{TAB}"
             Case 48 To 57
                If .SelStart = 0 Then .SelText = "("
                If .SelStart = 3 Then .SelText = ")" & Space(1)
                If .SelStart = 10 Then .SelText = "-"
             Case Else: KeyAscii = 0
          End Select
   End With
End Sub

Private Sub ToOnlyNull_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
       KeyAscii = 0
   End If
End Sub
Private Sub ToCurrency_Change()
   Dim I As Integer, t As String
   If Not IsNumeric(ToCurrency.Value) Then ToCurrency.Value = Empty
       With ToCurrency
           t = .Text
           I = Len(t) - .SelStart
           t = Replace(.Text, ",", "")
           If Len(t) < 3 Then t = String(3 - Len(t), "0") & t
           t = Mid(t, 1, Len(t) - 2) & "," & Mid(t, Len(t) - 1)
           t = Format(t, "R$#,##0.00")
           If .Text <> t Then .Text = t
           .SelStart = Len(t) - I
       End With
End Sub

