Operation =3
Name ="tblUnitSeriesSubunit"
Option =2
Where ="(((qrySearchAllSubs.[New Oracle Part #])=SubUnitExist()))"
Begin InputTables
    Name ="qrySearchAllSubs"
End
Begin OutputColumns
    Name ="Fiber Oracle item"
    Expression ="qrySearchAllSubs.[Fiber Oracle item]"
    Alias ="Expr1"
    Name ="Expr1"
    Expression ="[qrySearchAllSubs]![Item No]"
    Name ="OracleStatus"
    Expression ="qrySearchAllSubs.OracleStatus"
    Name ="New Oracle Part #"
    Expression ="qrySearchAllSubs.[New Oracle Part #]"
    Name ="Active"
    Expression ="qrySearchAllSubs.Active"
    Name ="Item No"
    Expression ="qrySearchAllSubs.[Item No]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa7887dad3d1cd94bb001233fa72b330b
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mid([Basic Product Construction_1]![Item No],11,2)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mid([Basic Product Construction_1]![Item No],10,1)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin_Subs.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySearchAllSubs.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin_Subs.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin_Subs.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySearchAllSubs.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin_Subs.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySearchAllSubs.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySearchAllSubs.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin_Subs.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySearchAllSubs.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18ba82b5b3270b40a5ebd335b3ed741e
        End
    End
End
Begin
    State =0
    Left =8
    Top =416
    Right =1117
    Bottom =985
    Left =-1
    Top =-1
    Right =1077
    Bottom =156
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =359
        Top =11
        Right =615
        Bottom =199
        Top =0
        Name ="qrySearchAllSubs"
        Name =""
    End
End
