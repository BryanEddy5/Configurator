dbMemo "SQL" ="SELECT [Basic Product Construction].Base, replace([Print Line 3],\"115304 \",\"\""
    ") AS PrintLine3, [Customer], [Listing Company], [New Oracle Part #] AS Oracle\015"
    "\012FROM [Basic Product Construction]\015\012WHERE [Basic Product Construction]."
    "[Active] <> 0 AND  (([Basic Product Construction].[New Oracle Part #])  not like"
    " \"*test*\") and  [Print Line 3] <> \"\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((Base Like \"UP*\"))"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf8bff3da48b7304bbce21a5e50d89543
End
dbMemo "OrderBy" ="Customer"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="PrintLine3"
        dbInteger "ColumnWidth" ="4980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x723be21c2647ff46afc116d25782396b
        End
    End
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
        dbText "Name" ="[Listing Company]"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Listing Company"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x29d13692899e2842b8af85a020364ed1
        End
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb1a7e200a8b68c42b1ab9a6e37843a14
        End
    End
End
