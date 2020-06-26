VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

Private Const DEFAULT_BLOCK_SIZE As Integer = 16

Private Type Class
   Buffer As String
   Length As String
End Type

Private This As Class

Private Sub Class_Initialize()
    This.Buffer = VBA.String(DEFAULT_BLOCK_SIZE, " ")
    This.Length = 0
End Sub

Public Property Get Length() As Integer
    Length = This.Length
End Property

Public Property Get Capacity() As Integer
    Capacity = Len(This.Buffer)
End Property

Public Function Append(ByVal Value As String) As StringBuilder
   Dim ValueLength As Integer
   
   ValueLength = Len(Value)
   ExpandIfNeeded (ValueLength)
   
   Mid(This.Buffer, This.Length + 1) = Value
   This.Length = This.Length + ValueLength
   Set Append = Me
End Function

Public Function Insert(ByVal Index As Integer, ByVal Value As String) As StringBuilder
   Dim ValueLength As Integer
   Dim Rest As String
   
   If Index < 0 Then
     Err.Raise 9
   ElseIf This.Length < Index Then
     Err.Raise 9
   End If
     
   ValueLength = Len(Value)
   Call ExpandIfNeeded(ValueLength)

   Rest = Mid(This.Buffer, Index + 1)
   Mid(This.Buffer, Index + 1) = Value
   Mid(This.Buffer, Index + ValueLength + 1) = Rest
   This.Length = This.Length + ValueLength
   
   Set Insert = Me
End Function

Private Sub ExpandIfNeeded(ByVal IncreaseSize)
   Dim NextLength As Integer
   Dim BlockSize As Integer
   
   NextLength = This.Length + IncreaseSize
   
   If Len(This.Buffer) < NextLength Then
      BlockSize = Max(IncreaseSize, This.Length)
      This.Buffer = This.Buffer + VBA.String(BlockSize, " ")
   End If
End Sub

Private Function Max(ByVal X As Integer, ByVal Y As Integer) As Integer
    Max = IIf(X > Y, X, Y)
End Function

Public Sub Clear()
    This.Buffer = ""
    This.Length = 0
End Sub

Public Function Equals(ByVal Value As StringBuilder) As Boolean
    If Value.ToString() <> Me.ToString() Then
        Equals = False
    ElseIf Value.Capacity <> Me.Capacity Then
        Equals = False
    Else
        Equals = True
    End If
End Function

Public Function ToString() As String
    ToString = Left(This.Buffer, This.Length)
End Function

Private Sub Class_Terminate()
    Call Clear
End Sub
