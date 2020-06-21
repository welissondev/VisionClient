VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DataList As Collection

Private Sub Class_Initialize()
   Set DataList = New Collection
End Sub

Public Sub Add(Value As Variant)
   DataList.Add Value
End Sub
Public Sub Remove(Index As Integer)
   DataList.Remove Index
End Sub

Public Function Item(Index As Integer) As Variant
   Item = DataList.Item(Index)
End Function

Public Function Count() As Integer
   Count = DataList.Count
End Function

Public Function GetValues() As String
   
   Dim i As Integer
   Dim Str As String
   
   For i = 1 To DataList.Count
      Str = Str & DataList.Item(i)
      If DataList.Count > 0 Then
         Str = Str & vbCrLf
      End If
   Next
   
   GetValues = Str
   
End Function



Private Sub Class_Terminate()
   Set DataList = Nothing
End Sub

