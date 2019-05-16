Option Compare Database
Option Explicit

Public Function FindTB() As Boolean

Dim cnn As ADODB.connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim strTB As String
Set cnn = CurrentProject.connection


strSQL = "SELECT  Left([New Oracle Part #],7) AS Oracle, Base, [Jacket Material], Count([New Oracle Part #]) AS [Count], Active "
strSQL = strSQL & "FROM [Basic Product Construction] "
strSQL = strSQL & "GROUP BY Left([New Oracle Part #],7), Base, [Jacket Material], Active "
strSQL = strSQL & "HAVING Base ='" & Forms!frmCreateItem!Base & "' AND Active <> 0 " 'AND (([Jacket Material])='" & Forms!frmCreateItem!Jacket & "')) "

strSQL = strSQL & "ORDER BY Count(Left([New Oracle Part #],7)) DESC"

rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic

If rst.RecordCount = 0 Then
    FindTB = False
    Exit Function
End If
'rst.MoveFirst

Forms!frmCreateItem!strTB = rst!Oracle


rst.Close
cnn.Close
Set cnn = Nothing
Set rst = Nothing



End Function