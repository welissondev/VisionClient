VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Generate(Optional Size As Integer = 8, Optional WithLetter As Boolean = False) As String
   
   Dim i As Integer

   For i = 1 To Size
      Select Case WithLetter
         Case Is = False
            Generate = Generate & Int((7 - 0 + 1) * Rnd + 0)
         Case Is = True
            Generate = Generate & Int((7 - 0 + 1) * Rnd + 0) & Chr(Int((22 * Rnd) + 1) + 64)
      End Select
   Next
   
End Function

