VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionAccessModel"
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
        Call This.Connection.Open(AppConfig.GetConnectionStrig)
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
    
    Call This.Command.execute(RecordsAffected)
    
    ExecuteNonQuery = RecordsAffected
End Function

Public Function ExecuteWithQuery(ByVal Sql As String) As ADODB.RecordSet

    This.Command.ActiveConnection = OpenConnection
    This.Command.CommandType = adCmdText
    This.Command.CommandText = Sql
    
    Set This.RecordSet = This.Command.execute()
    
    Set ExecuteWithQuery = This.RecordSet
    
End Function

Public Sub AddParameter(Name As String, Value As String, DataType As DataTypeEnum, Optional Size As Integer = 255)
    This.Command.Parameters.Append This.Command.CreateParameter(Name, DataType, adParamInput, Size, Value)
End Sub

Private Sub Class_Finalize()
   Set This.Connection = Nothing
   Set This.RecordSet = Nothing
   Set This.Command = Nothing
End Sub

