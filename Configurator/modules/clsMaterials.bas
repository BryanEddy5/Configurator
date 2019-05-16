Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Function LoadData(InvokeType As Integer) As Boolean

DoCmd.SetWarnings False

Dim rsMaterial As New ADODB.Recordset
Dim cnnMaterial As New ADODB.connection
Set cnnMaterial = CurrentProject.connection
Dim strSQL As String
Dim i As Integer
Dim mkey

strSQL = "Select * from tblMaterialBOMExplode"

rsMaterial.Open (strSQL), cnnMaterial, adOpenKeyset, adLockOptimistic


Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rs As New ADODB.Recordset
Dim sExec As String, OracleItem As String
Dim Quantity As Long
Dim sql1 As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
cnn.Open Core_SqlConnections.ConnectionString




sExec = " IF OBJECT_ID(N'tempdb..##ItemPassToSQL', N'U') IS NOT NULL "
sExec = sExec & " DROP TABLE ##ItemPassToSQL "
sExec = sExec & " CREATE TABLE ##ItemPassToSQL(OracleItem nvarchar(50),Quantity float,) "

cnn.Execute sExec

With rsMaterial

    Do Until .EOF
        
        OracleItem = Nz(!OracleItem)
        Quantity = Nz(!Quantity)
    
        sql1 = "insert into ##ItemPassToSQL values('" & OracleItem & "'," & Quantity & ")"
        
        'Debug.Print sql1
        cnn.Execute sql1
        .MoveNext
    Loop

End With




If InvokeType = 1 Then
    Call GetBuyMaterials
ElseIf InvokeType = 2 Then
    Call WhereUsed
Else
    MsgBox "Missing Invoke parameter", vbCritical
End If


''cn.Execute "drop table #myTemptable"
cnn.Close
Set cnn = Nothing




Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading data:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function
Private Function GetBuyMaterials() As Boolean

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim strStoredProcedure As String
Dim rstTable As New ADODB.Recordset
Dim cnnTable As New ADODB.connection



Set cnnTable = CurrentProject.connection

DoCmd.RunSQL "DELETE FROM tblGetBuyMaterials"



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
cnn.CommandTimeout = 600



strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[usp_GetBuyMaterials] "

rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic
'
rstTable.Open "tblGetBuyMaterials", cnnTable, adOpenKeyset, adLockOptimistic

DoCmd.OpenQuery "qryGetBuyMaterials_LOAD"


rst.Close
rstTable.Close
cnnTable.Close
cnn.Close



'
End Function


Private Function WhereUsed() As Boolean


Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim strStoredProcedure As String
Dim rstTable As New ADODB.Recordset
Dim cnnTable As New ADODB.connection
Dim start As Double
Dim Count%


On Error GoTo ErrorHandler

Set cnnTable = CurrentProject.connection

DoCmd.RunSQL "DELETE FROM tblGetWhereUsedMaterial"



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
cnn.CommandTimeout = 600



strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[usp_GetWhereUsedMaterials] "

rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

start = Timer
While Timer < start + 1
    DoEvents
Wend


'
'rstTable.Open "tblGetBuyMaterials", cnnTable, adOpenKeyset, adLockOptimistic

DoCmd.OpenQuery "qryGetWhereUsed_LOAD"
'


ErrorExit:
cnn.Close
Exit Function

ErrorHandler:
If Err.Number = 3146 And Count < 100 Then
    If Count < 10 Then
        start = Timer
        While Timer < start + 1
            DoEvents
        Wend
    End If
    Count = Count + 1
    Resume

Else
    MsgBox Err.Number & " " & Err.Description

End If
    Resume ErrorExit


'
End Function