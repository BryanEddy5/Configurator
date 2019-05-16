Operation =1
Option =2
Where ="((([Basic Product Construction].Active)=True) AND (([Basic Product Construction]"
    ".[New Oracle Part #]) Not Like \"*test*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Item No]"
    Alias ="Fiber"
    Expression ="Left([Fiber Oracle item],8)"
    Expression ="[Basic Product Construction].Customer"
    Alias ="Family"
    Expression ="Left([Item No],2)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x59a85b2eac282841ad37955c64384535
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Item No]"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fiber"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xefdf91bdd918e448b8b7078f33c101a9
        End
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Customer"
        dbInteger "ColumnWidth" ="4200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ce21305958dcd4caf72728fb3f89ba0
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-417
    Top =180
    Right =1420
    Bottom =659
    Left =-1
    Top =-1
    Right =1805
    Bottom =214
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
