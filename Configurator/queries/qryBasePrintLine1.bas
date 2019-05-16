dbMemo "SQL" ="SELECT [Print Line 1] AS PrintLine1, Trim([Customer]), [New Oracle Part #]\015\012"
    "FROM [Basic Product Construction]\015\012WHERE [Basic Product Construction].[Act"
    "ive] <> 0 AND  (([Basic Product Construction].[New Oracle Part #])  not like \"*"
    "test*\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((Expr1001 Like \"afl*\"))"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa57dfaef6f29d745a753d364bcfa0df4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="PrintLine1"
        dbInteger "ColumnWidth" ="3825"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x941ce0d4f6a9c445afc3e1f78d871f13
        End
    End
    Begin
        dbText "Name" ="Expr1001"
        dbInteger "ColumnWidth" ="3780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f6ba6b795b3a34fa51e17680fdb2ec2
        End
    End
    Begin
        dbText "Name" ="New Oracle Part #"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0dfdf7bbfb88ef40a1b8d892ed6dc406
        End
    End
End
