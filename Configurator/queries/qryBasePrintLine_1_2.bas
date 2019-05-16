dbMemo "SQL" ="SELECT [Print Line 1] AS PrintLine1, Trim([Customer]), [New Oracle Part #], [Pri"
    "nt Line 2] AS PrintLine2\015\012FROM [Basic Product Construction]\015\012WHERE ["
    "Basic Product Construction].[Active] = -1 AND  (([Basic Product Construction].[N"
    "ew Oracle Part #])  not like \"*test*\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf4506fad40a6c847a903a95344d5945d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="PrintLine1"
        dbInteger "ColumnWidth" ="7050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6cc39c8a19e1544e999b0f9ed1016fad
        End
    End
    Begin
        dbText "Name" ="Expr1001"
        dbInteger "ColumnWidth" ="3780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8cb3f115b8c2b438dde28fe32a113a3
        End
    End
    Begin
        dbText "Name" ="[New Oracle Part #]"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrintLine2"
        dbInteger "ColumnWidth" ="6735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b183d717f5bb54c9fcea291f61a61af
        End
    End
    Begin
        dbText "Name" ="New Oracle Part #"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdbab7d129acdd14a92c0c35b9c406faf
        End
    End
End
