Operation =1
Option =2
Where ="(((qryFiberOracle.Oracle) Not Like \"*TEST*\") AND ((qryFiberOracle.active)<>0))"
Begin InputTables
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
    Expression ="qryFiberOracle.Fiber2"
    Expression ="qryFiberOracle.PrintLine1"
    Expression ="qryFiberOracle.PrintLine2"
    Expression ="qryFiberOracle.PrintLine3"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.RibbonHighCure"
    Expression ="qryFiberOracle.PID"
    Expression ="qryFiberOracle.[Unit Series]"
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
    0x882e7515f06ea9498eb6b7c05443c581
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
        dbText "Name" ="qryFiberOracle.Fiber2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.RibbonHighCure"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ColorChip"
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
        dbText "Name" ="qryFiberOracle.Customer"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="qryFiberOracle.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Unit Series]"
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
    Bottom =316
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =501
        Top =19
        Right =845
        Bottom =348
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
End
