dbMemo "SQL" ="SELECT [Fiber Oracle Item], [Item No], OracleStatus, [New Oracle Part #], [Activ"
    "e], Base, [FiberType2]\015\012FROM [Basic Product Construction]\015\012\015\012U"
    "NION ALL SELECT [Fiber Oracle Item], [Item No], OracleStatus, [New Oracle Part #"
    "], [Active], Base, FiberType2\015\012FROM tblNewProductConstructionHold;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb6d90ee044e4ba4481b807162d32499d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="([qrySearchAllSubs].[New Oracle Part #] Like \"pe*\")"
dbMemo "OrderBy" ="[qrySearchAllSubs].[New Oracle Part #] DESC"
Begin
    Begin
        dbText "Name" ="[Active]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Fiber Oracle Item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Item No]"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OracleStatus"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ae10079cfa01345a5ca18dc4b331999
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b0dd87a9da83d49ba6cfdd0c7683d2e
        End
    End
    Begin
        dbText "Name" ="[FiberType2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fiber Oracle Item"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x14fbeeec251fcb4b9ca61520b833ea9b
        End
    End
    Begin
        dbText "Name" ="Item No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6229fdbcdda23b4c8c3a323ff97da82b
        End
    End
    Begin
        dbText "Name" ="New Oracle Part #"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xded25aff11b77548831375c9f47eebd7
        End
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x408e6becd8902a41aead08ef2574ce0d
        End
    End
    Begin
        dbText "Name" ="FiberType2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xea051b1d079a424f8a610f3cf95ed405
        End
    End
End
