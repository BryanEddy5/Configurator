Operation =1
Option =2
Where ="(((qryFiberOracle.Fiber) Like \"fbr00207*\") AND ((qryDesignCodeBreakdownAll.Fib"
    "erType)<>\"k\")) OR (((qryDesignCodeBreakdownAll.FiberType)=\"h\" And (qryDesign"
    "CodeBreakdownAll.FiberType)<>\"k\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="qryDesignCodeBreakdownAll"
    Name ="tblWOPExtract"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.OracleStatus"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryDesignCodeBreakdownAll.FiberType"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="tblWOPExtract.LotNumber"
    Expression ="tblWOPExtract.Locator"
    Alias ="Expr1"
    Expression ="tblWOPExtract.Qty"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="qryDesignCodeBreakdownAll"
    Expression ="qryFiberOracle.Oracle = qryDesignCodeBreakdownAll.Oracle"
    Flag =1
    LeftTable ="qryDesignCodeBreakdownAll"
    RightTable ="tblWOPExtract"
    Expression ="qryDesignCodeBreakdownAll.Oracle = tblWOPExtract.Item"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5e9a9df37348dc448b94f058c87ce247
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.OracleStatus"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdownAll.FiberType"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnWidth" ="2010"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="2595"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.Locator"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.[Item Number]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.LotNumber"
        dbInteger "ColumnWidth" ="3075"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Locator"
        dbInteger "ColumnWidth" ="1905"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Qty"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeac7bd732dca9946a909a7415f9b542e
        End
    End
End
Begin
    State =0
    Left =121
    Top =286
    Right =1887
    Bottom =926
    Left =-1
    Top =-1
    Right =1734
    Bottom =209
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =214
        Bottom =232
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =389
        Top =3
        Right =654
        Bottom =231
        Top =0
        Name ="qryDesignCodeBreakdownAll"
        Name =""
    End
    Begin
        Left =764
        Top =13
        Right =1017
        Bottom =217
        Top =0
        Name ="tblWOPExtract"
        Name =""
    End
End
