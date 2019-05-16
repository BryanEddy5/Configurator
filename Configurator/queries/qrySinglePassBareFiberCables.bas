Operation =1
Option =0
Where ="(((qryFiberOracle.Fiber) Like \"*FBR00424*\") AND ((qryFiberOracle.[Item No]) No"
    "t Like \"e*\" And (qryFiberOracle.[Item No]) Not Like \"t*\") AND ((qryItemNoJoi"
    "n.Location)=1) AND ((qryDesignCodeBreakdownAll.TBType)=\"B\"))"
Begin InputTables
    Name ="qryItemNoJoin"
    Name ="qryFiberOracle"
    Name ="qryDesignCodeBreakdownAll"
End
Begin OutputColumns
    Expression ="qryItemNoJoin.Oracle"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryFiberOracle.OracleStatus"
    Expression ="qryFiberOracle.Base"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryItemNoJoin.Location"
    Expression ="qryDesignCodeBreakdownAll.TBType"
    Expression ="qryDesignCodeBreakdownAll.FiberType"
End
Begin Joins
    LeftTable ="qryItemNoJoin"
    RightTable ="qryFiberOracle"
    Expression ="qryItemNoJoin.Oracle = qryFiberOracle.Oracle"
    Flag =1
    LeftTable ="qryFiberOracle"
    RightTable ="qryDesignCodeBreakdownAll"
    Expression ="qryFiberOracle.Oracle = qryDesignCodeBreakdownAll.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2d50598d13672b47a6858f730c99bbae
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Base"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="4860"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnOrder" ="2"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.OracleStatus"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin.Location"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdown.TBType"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdown.FiberType"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdownAll.TBType"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdownAll.FiberType"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =131
    Top =130
    Right =1240
    Bottom =759
    Left =-1
    Top =-1
    Right =1077
    Bottom =223
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =93
        Top =8
        Right =502
        Bottom =229
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
    Begin
        Left =557
        Top =4
        Right =914
        Bottom =242
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =976
        Top =6
        Right =1258
        Bottom =229
        Top =0
        Name ="qryDesignCodeBreakdownAll"
        Name =""
    End
End
