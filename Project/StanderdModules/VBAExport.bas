Attribute VB_Name = "VBAExport"
Option Explicit

Private Work As Workbook

Public Sub OverwriteExportedFile()

   Dim WindowsFormsPaste As String
   Dim StanderdModulePaste As String
   Dim ClassModulePaste As String
   Dim ComponentIndex As Long
   
   Set Work = ThisWorkbook
   
   WindowsFormsPaste = Work.Path & "\Project\WindowsForms\"
   StanderdModulePaste = Work.Path & "\Project\StanderdModules\"
   ClassModulePaste = Work.Path & "\Project\ClassModules\"
     
   Call DeleteFile(WindowsFormsPaste)
   Call DeleteFile(StanderdModulePaste)
   Call DeleteFile(ClassModulePaste)
   
   Call ExportComponents
   
End Sub

Public Sub ExportComponents()
    
   Dim WindowsFormsPaste As String
   Dim StanderdModulePaste As String
   Dim ClassModulePaste As String
   Dim ComponentIndex As Long
   
   Set Work = ThisWorkbook
   
   WindowsFormsPaste = Work.Path & "\Project\WindowsForms\"
   StanderdModulePaste = Work.Path & "\Project\StanderdModules\"
   ClassModulePaste = Work.Path & "\Project\ClassModules\"
           
   For ComponentIndex = 1 To Work.VBProject.VBComponents.Count
      
      DoEvents
      
      With Work.VBProject.VBComponents(ComponentIndex)
         
         If .Name <> "VBATest" Then
         
               Select Case .Type
               
                    Case vbext_ct_MSForm
                         Call ExportVBComponent(Work.VBProject.VBComponents(.Name), WindowsFormsPaste)
                         
                    Case vbext_ct_StdModule
                         Call ExportVBComponent(Work.VBProject.VBComponents(.Name), StanderdModulePaste)
                         
                    Case vbext_ct_ClassModule
                         Call ExportVBComponent(Work.VBProject.VBComponents(.Name), ClassModulePaste)
                         
                  End Select
            
         End If
         
      End With
      
   Next
   
   Debug.Print "--------------------------------------"
   Debug.Print "Sucessfull " & Date & " " & Time
   Debug.Print ""
End Sub

Public Sub ExportWorkSheet()
   
   Dim Sheet As FileSystemObject
   Dim SourceDirectory, DestinationDirectory As String
   Dim Override As Boolean

   Set Work = ThisWorkbook
   Set Sheet = New FileSystemObject
   
   SourceDirectory = Work.Path + "\" + Work.Name
   DestinationDirectory = Work.Path + "\Project\" + Work.Name
   Override = True
   
   Call Sheet.CopyFile(SourceDirectory, DestinationDirectory, Override)
   
   Set Work = Nothing
   Set Sheet = Nothing

   Debug.Print "--------------------------------------"
   Debug.Print "Sucessfull " & Date & " " & Time
   Debug.Print ""
   
End Sub

Private Sub DeleteFile(Path As String)

    Dim FileList() As String
    Dim Index As Long
    
    Set Work = ThisWorkbook
    
    FileList = ListFiles(Path)
    
    For Index = 0 To UBound(FileList)
       
       DoEvents
       
       If FileList(Index) <> Empty Then
             Call Kill(Path & FileList(Index))
       End If
       
    Next

End Sub


Private Function ListFiles(ByVal Path As String) As String()
       
       On Error GoTo Excption
       
          Dim FileSystem As New FileSystemObject
          Dim ArrayList() As String
          Dim RootPaste As Folder
          Dim FileStream As File
          Dim FileIndex As Long
       
          ReDim ArrayList(0) As String
          
          If FileSystem.FolderExists(Path) Then
              
              Set RootPaste = FileSystem.GetFolder(Path)
       
              For Each FileStream In RootPaste.Files
                  FileIndex = IIf(ArrayList(0) = "", 0, FileIndex + 1)
                  ReDim Preserve ArrayList(FileIndex) As String
                  ArrayList(FileIndex) = FileStream.Name
              Next
              
          End If
       
          ListFiles = ArrayList
    
Excption:

    Set FileSystem = Nothing
    Set RootPaste = Nothing
    Set FileStream = Nothing
    
End Function


Private Function ExportVBComponent(Conponent As VBIDE.VBComponent, FolderName As String, Optional FileName As String, _
Optional OverwriteExisting As Boolean = True) As Boolean

       Dim Extension As String
       Dim ConponentName As String
       
       Extension = GetFileExtension(Conponent:=Conponent)
       
       If Trim(FileName) = vbNullString Then
           ConponentName = Conponent.Name & Extension
       Else
           ConponentName = FileName
           If InStr(1, ConponentName, ".", vbBinaryCompare) = 0 Then
               ConponentName = ConponentName & Extension
           End If
       End If
       
       If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
           ConponentName = FolderName & ConponentName
       Else
           ConponentName = FolderName & "\" & ConponentName
       End If
       
       If Dir(ConponentName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
           If OverwriteExisting = True Then
               Kill ConponentName
           Else
               ExportVBComponent = False
               Exit Function
           End If
       End If
       
       Conponent.Export FileName:=ConponentName
       ExportVBComponent = True
    
End Function
    
    
Private Function GetFileExtension(Conponent As VBIDE.VBComponent) As String

   Select Case Conponent.Type
         Case vbext_ct_ClassModule
             GetFileExtension = ".cls"
         Case vbext_ct_Document
             GetFileExtension = ".cls"
         Case vbext_ct_MSForm
             GetFileExtension = ".frm"
         Case vbext_ct_StdModule
             GetFileExtension = ".bas"
         Case Else
             GetFileExtension = ".bas"
   End Select

End Function

