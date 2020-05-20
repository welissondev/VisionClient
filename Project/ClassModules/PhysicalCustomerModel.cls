VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhysicalCustomerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Insert() As Boolean
       
    Dim DataBase As ConnectionAccessModel
    Dim Query As String
    
    Query = "INSERT INTO PhysicalCostumer(InternalCode, YourName, IndentyCard, SocialSecurity)Values(?,?,?,?)"
    
    Set DataBase = New ConnectionAccessModel
    
    With DataBase
        Insert = .Command.ExecuteNonQuery(.Connection, Query, "568", "Maria Silva", "6656052245", "5565330666")
    End With
    
    Set DataBase = Nothing
    
End Function


Public Sub ListData(DataList As Object)

    Dim DataBase As ConnectionAccessModel
    Dim DataListIndex As Integer
    Dim Query As String
    
    Query = "SELECT * FROM PhysicalCostumer"
    
    Set DataBase = New ConnectionAccessModel
        Call DataBase.ExecuteWithQuery(Query)
    
    Do Until DataBase.RecordSet.EOF
        With DataList
            .ColumnWidths = "00;71;180;90;80;150;00"
            .AddItem
                .List(DataListIndex, 0) = "" & DataBase.RecordSet.Fields("Id").Value
                .List(DataListIndex, 1) = "" & DataBase.RecordSet.Fields("InternalCode").Value
                .List(DataListIndex, 2) = "" & DataBase.RecordSet.Fields("YourName").Value
                .List(DataListIndex, 3) = "" & DataBase.RecordSet.Fields("IndentyCard").Value
                .List(DataListIndex, 4) = "" & DataBase.RecordSet.Fields("SocialSecurity").Value
                .List(DataListIndex, 5) = "" & DataBase.RecordSet.Fields("Email").Value
            DataListIndex = DataListIndex + 1
       End With
       DataBase.RecordSet.MoveNext
    Loop

    Set DataBase = Nothing
     
End Sub


