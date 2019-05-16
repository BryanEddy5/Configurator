Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Like \"pt*\" Or ([Basic Pro"
    "duct Construction].[New Oracle Part #]) Like \"rbn*\") AND (([Basic Product Cons"
    "truction].Active)<>0))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Alias ="ItemNumberSuffix"
    Expression ="Mid([New Oracle Part #],9,2)"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[EZ Strip]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xb4e224004c122500d08f8001ffffffff
End
Begin
    Begin
        dbText "Name" ="ItemNumberSuffix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0afcedf464987f4485626ccbad680c96
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbInteger "ColumnWidth" ="3600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =182
    Right =1190
    Bottom =713
    Left =-1
    Top =-1
    Right =1077
    Bottom =293
    Left =1136
    Top =0
    ColumnsShown =539
    Begin
        Left =-1098
        Top =6
        Right =-885
        Bottom =319
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
