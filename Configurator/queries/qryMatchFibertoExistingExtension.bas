Operation =1
Option =2
Where ="(((Mid([Forms]![frmCreateItem]![ItemNo],14,30)) Like \"*\" & [FiberExtension] & "
    "\"*\"))"
Begin InputTables
    Name ="tbl_Fibers"
    Name ="tblFiberExtensions"
    Name ="qryFiberOracle"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryFiberOracle.Color"
    Expression ="qryFiberOracle.Jacket"
    Expression ="qryFiberOracle.Base"
    Expression ="qryFiberOracle.ColorChip"
    Expression ="qryFiberOracle.Customer"
    Expression ="qryFiberOracle.PrintLine4"
    Expression ="qryFiberOracle.active"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.Fiber2"
    Expression ="tbl_Fibers.BIF"
    Expression ="qryFiberOracle.PrintLine1"
    Expression ="qryFiberOracle.PrintLine3"
    Expression ="qryFiberOracle.PrintLine2"
    Expression ="tblFiberExtensions.FiberExtension"
    Expression ="qryFiberOracle.PID"
    Expression ="qryFiberOracle.RibbonHighCure"
    Expression ="qryFiberOracle.[Unit Series]"
End
Begin Joins
    LeftTable ="tbl_Fibers"
    RightTable ="qryFiberOracle"
    Expression ="tbl_Fibers.Fiber = qryFiberOracle.Fiber"
    Flag =1
    LeftTable ="qryFiberOracle"
    RightTable ="tblFiberExtensions"
    Expression ="qryFiberOracle.Fiber = tblFiberExtensions.Fiber"
    Flag =1
End
Begin OrderBy
    Expression ="qryFiberOracle.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3d75724cf5767e4ca47c6505f778f7a2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblFiberExtensions.FiberExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.BIF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.RibbonHighCure"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =164
    Right =1373
    Bottom =887
    Left =-1
    Top =-1
    Right =1260
    Bottom =209
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =99
        Top =13
        Right =401
        Bottom =247
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =1062
        Top =12
        Right =1206
        Bottom =156
        Top =0
        Name ="tblFiberExtensions"
        Name =""
    End
    Begin
        Left =479
        Top =9
        Right =803
        Bottom =216
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
End
