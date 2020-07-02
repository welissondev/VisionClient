VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Class
    Connection As ADODB.Connection
    Recordset As ADODB.Recordset
    Command As ADODB.Command
End Type

Private This As Class

Private Sub Class_Initialize()
    Set This.Connection = New ADODB.Connection
    Set This.Command = New ADODB.Command
    Set This.Recordset = New ADODB.Recordset
End Sub
Private Sub Class_Terminate()
   If This.Connection.State = True Then
       This.Connection.Close
   End If
   If This.Recordset.State = True Then
      This.Recordset.Close
   End If
   
   Set This.Recordset = Nothing
   Set This.Connection = Nothing
   Set This.Command = Nothing
End Sub

Public Property Get Connection() As ADODB.Connection
    Set Connection = This.Connection
End Property

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = This.Recordset
End Property
Public Property Let Recordset(Recordset As ADODB.Recordset)
    Set This.Recordset = Recordset
End Property

Public Function OpenConnection() As ADODB.Connection
   If This.Connection.State = False Then
       Call This.Connection.Open(SysPropertys.ConnectionString)
   End If
   Set OpenConnection = This.Connection
End Function

Public Sub CloseConnection()
    If This.Connection.State = True Then This.Connection.Close
End Sub

Public Function ExecuteNonQuery(ByVal Sql As String) As Boolean
    
    Dim RecordsAffected As Integer
    
    This.Command.ActiveConnection = OpenConnection
    This.Command.CommandType = adCmdText
    This.Command.CommandText = Sql
    
    Call This.Command.Execute(RecordsAffected)
   
    ExecuteNonQuery = RecordsAffected
    
End Function

Public Function ExecuteWithQuery(ByVal Sql As String) As ADODB.Recordset

    This.Command.ActiveConnection = OpenConnection
    This.Command.CommandType = adCmdText
    This.Command.CommandText = Sql
    This.Connection.CursorLocation = adUseClient
    
    Set This.Recordset = This.Command.Execute
    Set ExecuteWithQuery = This.Recordset
    
End Function

Public Sub AddParameter(Name As String, Value As Variant, Optional DataType As DataTypeEnum = adEmpty, _
Optional ParameterDirection As ParameterDirectionEnum = adParamInput, Optional Size As Integer = 255)
   This.Command.Parameters.Append This.Command.CreateParameter(Name, DataType, _
   ParameterDirection, Size, Value)
End Sub


Public Sub ClearParameter()
   
   Dim I As Integer
   
   If This.Command.Parameters.Count > 0 Then
      For I = 0 To This.Command.Parameters.Count - 1
         Call This.Command.Parameters.Delete(I)
      Next
   End If
   
End Sub

Public Property Get Field(Name As String) As Variant
    Field = Recordset.Fields(Name).Value
End Property
