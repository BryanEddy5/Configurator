Operation =1
Option =2
Where ="(((qryFiberOracle.Oracle) Not Like \"fb*\" And ((qryFiberOracle.Oracle) Like \"*"
    "-01*\" Or (qryFiberOracle.Oracle) Like \"*-00*\")) AND ((tbl_Fibers.Fiber) Like "
    "\"*161*\" Or (tbl_Fibers.Fiber) Like \"*119*\" Or (tbl_Fibers.Fiber) Like \"*131"
    "*\") AND ((tblAFLPRD_INVSysItemCost_CAB.Item_Status) Not Like \"*obs*\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tbl_Fibers"
    Name ="tblAFLPRD_INVSysItemCost_CAB"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Alias ="Expr1"
    Expression ="tbl_Fibers.Fiber"
    Alias ="Expr2"
    Expression ="tbl_Fibers.Mode"
    Expression ="qryFiberOracle.[Item No]"
    Alias ="Expr3"
    Expression ="tbl_Fibers.BIF"
    Alias ="Expr4"
    Expression ="tbl_Fibers.[OM#]"
    Expression ="tblAFLPRD_INVSysItemCost_CAB.Item_Status"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="tblAFLPRD_INVSysItemCost_CAB"
    Expression ="qryFiberOracle.Oracle = tblAFLPRD_INVSysItemCost_CAB.ItemNumber"
    Flag =1
    LeftTable ="tbl_Fibers"
    RightTable ="qryFiberOracle"
    Expression ="tbl_Fibers.Fiber = qryFiberOracle.Fiber"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Fibers.Fiber"
    Flag =0
    Expression ="qryFiberOracle.[Item No]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf834c59c1a7981498f72ff1abc7c94b8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.BIF"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Fiber"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Mode"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.[OM#]"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.SpecialNotes"
        dbInteger "ColumnWidth" ="1695"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="3885"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.Item_Status"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="4365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbInteger "ColumnWidth" ="5595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6dc7a72e35d09346a324f64a9607d995
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ec988f0ad64374097b11afd34b4177c
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb97d6bdc75e51b449264a63f53daa861
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd551e00619dfb8459ed787f38746c9d5
        End
    End
End
Begin
    State =0
    Left =47
    Top =124
    Right =1339
    Bottom =856
    Left =-1
    Top =-1
    Right =1260
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =215
        Bottom =199
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =376
        Top =60
        Right =594
        Bottom =275
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =669
        Top =6
        Right =916
        Bottom =214
        Top =0
        Name ="tblAFLPRD_INVSysItemCost_CAB"
        Name =""
    End
End
