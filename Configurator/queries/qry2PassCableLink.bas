Operation =1
Option =0
Begin InputTables
    Name ="dbo_CableUnion"
    Name ="Unit Series"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="dbo_CableUnion.Oracle"
    Expression ="dbo_CableUnion.Base"
    Expression ="[Unit Series].[Unit Color Series]"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].Reason"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Expression ="[Basic Product Construction].FiberType2"
    Expression ="[Unit Series].[Unit ID]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Unit Series].[Fiber Code]"
    Expression ="[Basic Product Construction].OracleStatus"
End
Begin Joins
    LeftTable ="dbo_CableUnion"
    RightTable ="Unit Series"
    Expression ="dbo_CableUnion.[Unit Series] = [Unit Series].[Unit Color Series]"
    Flag =1
    LeftTable ="Unit Series"
    RightTable ="Basic Product Construction"
    Expression ="[Unit Series].OracleItem = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
Begin OrderBy
    Expression ="[Unit Series].[Unit ID]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x50d16b18be78b043972a5f2e9b32a5f1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Unit Series].[Unit Color Series]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbInteger "ColumnWidth" ="3090"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Unit Series].[Fiber Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Unit Series].[Unit ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Base"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =74
    Right =1767
    Bottom =831
    Left =-1
    Top =-1
    Right =1703
    Bottom =298
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =378
        Top =21
        Right =600
        Bottom =264
        Top =0
        Name ="dbo_CableUnion"
        Name =""
    End
    Begin
        Left =776
        Top =22
        Right =955
        Bottom =200
        Top =0
        Name ="Unit Series"
        Name =""
    End
    Begin
        Left =1073
        Top =21
        Right =1317
        Bottom =232
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
