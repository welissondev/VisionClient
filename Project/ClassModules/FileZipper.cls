VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As Variant
Public SourcePath As Variant
Public DestinationPath As Variant

Public Function ZipFile() As Boolean
    
    On Error GoTo ErrorHamdle
        
        Dim App As Object
        
        Call CreateZipEmpty(DestinationPath & "\" & FileName & ".zip")
        
        Set App = CreateObject("Shell.Application")
            App.Namespace(DestinationPath & "\" & FileName & ".zip").CopyHere App.Namespace(SourcePath & "\").Items
            
        Do Until App.Namespace(DestinationPath & "\" & FileName & ".zip").Items.Count = App.Namespace(SourcePath & "\").Items.Count
            DoEvents
            Application.Wait (Now + TimeValue("0:00:03"))
        Loop
        
        ZipFile = True
    
      Exit Function

ErrorHamdle:

    ZipFile = False
    
End Function

Public Sub CreateZipEmpty(Path As Variant)
    If Len(Dir(Path)) > 0 Then Kill Path
      Open Path For Output As #1
      Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub
