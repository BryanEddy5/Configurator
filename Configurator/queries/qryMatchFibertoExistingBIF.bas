Operation =1
Option =2
Where ="(((qryFiberOracle.Active)<>0) AND ((tbl_Fibers.BIF)=\"yes\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tbl_Fibers"
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
    Expression ="qryFiberOracle.Active"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Fiber2"
    Alias ="BIF_"
    Expression ="tbl_Fibers.BIF"
    Expression ="qryFiberOracle.PrintLine1"
    Expression ="qryFiberOracle.PrintLine2"
    Expression ="qryFiberOracle.PrintLine3"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.PID"
    Expression ="qryFiberOracle.[Unit Series]"
    Expression ="qryFiberOracle.RibbonHighCure"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="tbl_Fibers"
    Expression ="qryFiberOracle.Fiber = tbl_Fibers.Fiber"
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
    0xbff2912b5e47ec4ebb42610c92c5983b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Base"
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
        dbText "Name" ="qryFiberOracle.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BIF_"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x84909ce5d705974e8fc16e8a092673f4
        End
    End
    Begin
        dbText "Name" ="qryFiberOracle.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine3"
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
        dbText "Name" ="qryFiberOracle.PID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.RibbonHighCure"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =265
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =207
        Top =16
        Right =439
        Bottom =267
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =617
        Top =17
        Right =887
        Bottom =197
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
