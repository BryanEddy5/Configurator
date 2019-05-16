Option Compare Database
Option Explicit

'Class developed to remove boilerplate code for retrieving, inserting, and updating recordsets from SQL db

Public Function GetAdodbRecordset(SqlQuery As String) As ADODB.Recordset

    Dim rst As New ADODB.Recordset

    rst.Open (SqlQuery), GetOpenSqlConnection(), adOpenStatic, adLockOptimistic
        
    Set GetAdodbRecordset = rst


End Function


Public Sub ExecuteSqlProcedureNoReturnOutput(SqlProcedure As String, ParamArray arg())

    Dim CM As New ADODB.Command
    Dim i As Integer
    Dim procDecorator As String
    
    procDecorator = FormatString("DECLARE @return_value INT EXEC  @return_value = {0} SELECT 'Return Value' = @return_value", SqlProcedure)
    
    With CM
        .ActiveConnection = GetOpenSqlConnection()
        .CommandType = adCmdStoredProc
        .CommandText = SqlProcedure
        .Parameters.Refresh
        For i = 0 To UBound(arg)
            CM(i + 1) = arg(i)
        Next i
        .Execute
    End With


End Sub

Public Function GetOpenSqlConnection() As ADODB.connection

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim provstr As String
    Dim path As String, separator As String
    
    Set cnn = New ADODB.connection
    cnn.Provider = "sqloledb"
    
    provstr = Core_SqlConnections.ConnectionString
    
    If cnn.State <> 1 Then cnn.Open provstr
    
    Set GetOpenSqlConnection = cnn
    

End Function

Public Function AddNewRecord(Table As String, Records As Collection)

Dim rs As New ADODB.Recordset

Set rs = Core_DataAccess.GetAdodbRecordset(Table)
    
    With rs
        .AddNew
            PassRecordValues Records, rs
        .Update
        .Close
        
    End With
    Set rs = Nothing

End Function

Private Sub PassRecordValues(NewRecords As Collection, OldRecords As ADODB.Recordset)

Dim f, k As ADODB.Field

For Each f In NewRecords

    For Each k In OldRecords
        If f.name = k.name Then
            k.value = f.value
        End If
    
    Next k
    
Next f

    
End Sub