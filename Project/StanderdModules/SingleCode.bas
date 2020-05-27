Attribute VB_Name = "SingleCode"
Option Explicit

Public Function GetCode(Optional Size As Integer = 8) As String
   
   Dim Count As Integer
   Dim ToString As String
      
   For Count = 1 To Size
      ToString = ToString & Int((7 - 0 + 1) * Rnd + 0)
   Next
       
   GetCode = ToString
                                 
End Function

