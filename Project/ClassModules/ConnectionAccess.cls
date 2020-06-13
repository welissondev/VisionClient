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
    RecordSet As ADODB.RecordSet
    Command As ADODB.Command
End Type

Private This As Class

Public Sub Class_Initialize()
    Set This.Connection = New ADODB.Connection
    Set This.Command = New ADODB.Command
    Set This.RecordSet = New ADODB.RecordSet
End Sub
Private Sub Class_Terminate()
   If This.Connection.State = True Then
       This.Connection.Close
   End If
   If This.RecordSet.State = True Then
      This.RecordSet.Close
   End If
   
   Set This.RecordSet = Nothing
   Set This.Connection = Nothing
   Set This.Command = Nothing
End Sub

Public Property Get Connection() As ADODB.Connection
    Set Connection = This.Connection
End Property

Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = This.RecordSet
End Property
Public Property Let RecordSet(RecordSet As ADODB.RecordSet)
    Set This.RecordSet = RecordSet
End Property

Public Function OpenConnection() As ADODB.Connection
   If This.Connection.State = False Then
       Call This.Connection.Open(AppSettings.ConnectionString)
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

Public Function ExecuteWithQuery(ByVal Sql As String) As ADODB.RecordSet

    This.Command.ActiveConnection = OpenConnection
    This.Command.CommandType = adCmdText
    This.Command.CommandText = Sql
    This.Connection.CursorLocation = adUseClient
    
    Set This.RecordSet = This.Command.Execute
    Set ExecuteWithQuery = This.RecordSet
    
End Function

Public Sub AddParameter(Name As String, Value As Variant, Optional DataType As DataTypeEnum = adEmpty, _
Optional ParameterDirection As ParameterDirectionEnum = adParamInput, Optional Size As Integer = 255)
   This.Command.Parameters.Append This.Command.CreateParameter(Name, DataType, _
   ParameterDirection, Size, Value)
End Sub
Public Sub ClearParameter()
   
   Dim i As Integer
   
   If This.Command.Parameters.Count > 0 Then
      For i = 0 To This.Command.Parameters.Count - 1
         Call This.Command.Parameters.Delete(i)
      Next
   End If
   
End Sub

