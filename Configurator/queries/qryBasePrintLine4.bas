dbMemo "SQL" ="SELECT [Print Line 4] AS PrintLine4, [Customer]\015\012FROM [Basic Product Const"
    "ruction]\015\012WHERE (([Basic Product Construction].[New Oracle Part #])  not l"
    "ike \"*test*\")\015\012ORDER BY [Basic Product Construction].Customer;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((Base Like \"UP*\"))"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "OrderBy" ="Customer"
dbBinary "GUID" = Begin
    0x4cbd8a77f9cc97469b1e1bc52bdc09a6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Customer]"
        dbInteger "ColumnWidth" ="3060"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe01dcc90a556974b919cd449a33c3c44
        End
    End
    Begin
        dbText "Name" ="PrintLine4"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x717e5f04bd03e64c9e52871268e6293f
        End
    End
End
