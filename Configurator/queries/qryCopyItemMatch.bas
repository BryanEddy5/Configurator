Operation =1
Option =16
RowCount ="1"
Where ="((([Basic Product Construction].Active)<>0) AND (([Basic Product Construction].["
    "New Oracle Part #]) Like \"P*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
Begin OrderBy
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6e023d523015434cbd59f905ca74b56b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Active"
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
End
Begin
    State =0
    Left =70
    Top =293
    Right =1557
    Bottom =830
    Left =-1
    Top =-1
    Right =1455
    Bottom =182
    Left =18
    Top =0
    ColumnsShown =539
    Begin
        Left =498
        Top =18
        Right =785
        Bottom =234
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
