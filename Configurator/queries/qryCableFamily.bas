dbMemo "SQL" ="SELECT left([Item No],2) AS CableFamily, Active\015\012FROM [Basic Product Const"
    "ruction]\015\012WHERE [Basic Product Construction].[New Oracle Part #] not like "
    "\"*test*\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5c2820ab295e6a4d9c9e02ec089f9a8b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CableFamily"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x944c12e31413cb41b85b8f8574de067a
        End
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x583dc13f5ac7254592b855478c182098
        End
    End
End
