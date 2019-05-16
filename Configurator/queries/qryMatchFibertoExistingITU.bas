Operation =1
Option =2
Where ="(((qryFiberOracle.active)<>0) And ((tbl_Fibers.ITU) Like \"*\" & Right(Forms!frm"
    "CreateItem!ItemNo,2) & \"*\"))"
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
    Expression ="qryFiberOracle.active"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.Fiber2"
    Expression ="tbl_Fibers.BIF"
    Expression ="tbl_Fibers.ITU"
    Expression ="qryFiberOracle.PrintLine1"
    Expression ="qryFiberOracle.PrintLine3"
    Expression ="qryFiberOracle.PrintLine2"
    Expression ="qryFiberOracle.[Unit Series]"
    Expression ="qryFiberOracle.RibbonHighCure"
    Expression ="qryFiberOracle.PID"
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
    0x442da77860369e40b9c9c05773cb0624
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
        dbText "Name" ="qryFiberOracle.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine1"
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
        dbText "Name" ="qryFiberOracle.active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
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
    Begin
        dbText "Name" ="qryFiberOracle.PID"
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
    Bottom =300
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =501
        Top =19
        Right =747
        Bottom =335
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =818
        Top =8
        Right =1088
        Bottom =186
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
