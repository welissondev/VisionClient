Attribute VB_Name = "SysMethods"
Option Explicit

Public Sub OpenPageWeb(Url As String)
    ActiveWorkbook.FollowHyperlink Address:=Url, NewWindow:=True, AddHistory:=True
End Sub

Public Sub IndentyDataTable(TableName As String, Optional Indent As Integer = 1, Optional SelectRange As String = "A1")

   Application.GoTo Reference:=TableName
      With Selection
         .HorizontalAlignment = xlGeneral
         .VerticalAlignment = xlCenter
         .InsertIndent Indent
      End With
   Range(SelectRange).Select
   
End Sub

Public Sub ClearTableContents(TableName As String, Optional SelectRange As String = "A1")

   Application.GoTo Reference:=TableName
   Selection.ClearContents
   Range(SelectRange).Select
   
End Sub

Public Sub ProtectSheet(Sheet As Worksheet, PassWord As String)
   Sheet.Protect PassWord:=PassWord
End Sub
Public Sub UnprotectSheet(Sheet As Worksheet, PassWord As String)
   Sheet.Unprotect PassWord:=PassWord
End Sub

Public Sub ActiveWorkbookRefreshAll()
   ActiveWorkbook.RefreshAll
End Sub

Public Sub SubmitException()

    Dim FileName As String
    Dim TextFile As Object
    Dim CurrenLine As String
    
    With New FileSystemObject
    
        FileName = SysDirectorys.PathAppLog & "\Error.txt"
            
        Set TextFile = .OpenTextFile(FileName, 1, True)
              While Not TextFile.AtEndOfStream
                  DoEvents
                  CurrenLine = CurrenLine & TextFile.ReadLine & vbCrLf
              Wend
          TextFile.Close
        
        Set TextFile = .OpenTextFile(FileName, 2, True)
          With TextFile
              .WriteLine "Error:" & Err.Number & Space(1) & "Description:" & Err.Description & vbCrLf & vbCrLf & CurrenLine
              .Close
          End With
              
    End With
     
    FormExceptionErrorNotifier.Show
    
End Sub

Public Sub DefineUserFormStyle(Form As MSForms.UserForm, Optional ByVal MainColor As Variant = 13408512)

   On Error GoTo Exception
   
      Dim Control As MSForms.Control
   
      Dim TextBorderColor, TextBackColor, TextForeColor, TextBorderStyle, TextFont As Variant
         TextFont = "Arial"
         TextBackColor = &H80000004
         TextBorderColor = &H80000016
         TextForeColor = 1842204
         TextBorderStyle = 1
      
      Dim LabelForeColor, LabelFont As Variant
         LabelForeColor = MainColor
         LabelFont = "Calibri"
         
      Dim FrameBackColor, FrameForeColor, FrameBorderColor, FrameFont As Variant
         FrameBorderColor = 14540253
         FrameForeColor = &H996600
         FrameBackColor = &HFFFFFF
         FrameFont = "Calibri"
      
      Form.Font = "Calibri"
      Form.BackColor = MainColor
      
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
                  If Control.Tag <> "Title" Then
                      .ForeColor = LabelForeColor
                      With .Font
                         .Name = LabelFont
                         .Size = 11
                      End With
                  End If
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
                  .ForeColor = &H996600
                  With .Font
                     .Name = TextFont
                     .Size = 10
                  End With
               End With
         End Select
      Next
      
   Exit Sub
   
Exception:
   
   Call SysMethods.SubmitException
   
End Sub

