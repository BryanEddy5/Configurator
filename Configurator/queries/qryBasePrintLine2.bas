dbMemo "SQL" ="SELECT [Basic Product Construction].Base, [Print Line 2] AS PrintLine2, [Custome"
    "r], [Fiber Oracle Item] AS Fiber, LEFT([Item No],5) AS CableFiber, [New Oracle P"
    "art #] AS Oracle\015\012FROM [Basic Product Construction]\015\012WHERE [Basic Pr"
    "oduct Construction].[Active] <> 0 AND  (([Basic Product Construction].[New Oracl"
    "e Part #])  not like \"*test*\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2f7bb6a6e7301e4d89f8a69ea2c41bb2
End
dbMemo "Filter" ="((qryBasePrintLine2.Fiber=\"FBR00397\"))"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="PrintLine2"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd6704ffc8903634682cfac15c58db555
        End
    End
    Begin
        dbText "Name" ="[Customer]"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fiber"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3448e06ef6c0674db695fddf48b9b831
        End
    End
    Begin
        dbText "Name" ="CableFiber"
        dbInteger "ColumnWidth" ="3645"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6497dc5429cac94589cffd5952cb3c4e
        End
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe34eb27aadf11744b2cc8916334a6c77
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x57d8a77e855676478b7951d674acb041
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x294a6a2bc91e6f4b999c73c81dd8950f
        End
    End
End
