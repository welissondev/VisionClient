VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private List As Collection

Private Sub Class_Initialize()
   Set List = New Collection
End Sub

Public Sub Add(Value As Variant)
   List.Add Value
End Sub
Public Sub Remove(Index As Integer)
   List.Remove Index
End Sub

Public Function Item(Index As Integer) As Variant
   Item = List.Item(Index)
End Function

Public Function Count() As Integer
   Count = List.Count
End Function

Private Sub Class_Terminate()
   Set List = Nothing
End Sub

