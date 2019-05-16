dbMemo "SQL" ="IF OBJECT_ID(N'tempdb..#ItemPassToSQL', N'U') IS NOT NULL \015\012DROP TABLE #It"
    "emPassToSQL\015\012\015\012CREATE TABLE #ItemPassToSQL(\015\012Item nvarchar(50)"
    ",\015\012Quantity float,\015\012)"
dbMemo "Connect" ="ODBC;DRIVER=SQL Server;SERVER=NAASPB-PRD04\\SQL2014;UID=eddybc;Trusted_Connectio"
    "n=Yes;DATABASE=Premise"
dbBoolean "ReturnsRecords" ="0"
dbInteger "ODBCTimeout" ="60"
dbBoolean "LogMessages" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x22db9fa8724ed4409e71c2b51bee0d18
End
Begin
End
