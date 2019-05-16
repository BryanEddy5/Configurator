Option Compare Database
Option Explicit


Function getBurnFileNumber(CustomerID As Integer, DocumentID As Integer) As String
    Dim rst As New ADODB.Recordset
    Dim cnn As New ADODB.connection
    Dim Procedure As String
    
    
    cnn.Provider = "sqloledb"
    cnn.Open Core_SqlConnections.ConnectionString
    
    Procedure = "DECLARE @return_value int, " & _
                        "@Enumber nvarchar(50) " & _
                "EXEC    @return_value = [dbo].[usp_GetBurnListingEnumber] @CustomerID = " & CustomerID & ", " & _
                "        @DocumentID = " & DocumentID & ", " & _
                "        @Enumber = @Enumber OUTPUT " & _
                "SELECT  @Enumber as N'Enumber' "

    
    rst.Open (Procedure), cnn, adOpenStatic, adLockReadOnly
    
    
    If rst.EOF = False Then
    
        getBurnFileNumber = rst!ENumber

    End If
    


End Function