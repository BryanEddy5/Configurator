Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Like Left([Forms]![frmFiber"
    "MatchExisting]![Oracle],7) & \"*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdce31b0d8db8eb4aa5293382292b3d0b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Basic Product Construction.Customer Part#"
        dbInteger "ColumnWidth" ="3600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Type (base)"
        dbInteger "ColumnWidth" ="2385"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 1"
        dbInteger "ColumnWidth" ="5490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 2"
        dbInteger "ColumnWidth" ="5895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 3"
        dbInteger "ColumnWidth" ="4440"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
