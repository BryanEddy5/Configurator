Operation =4
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Not Like \"*test*\") AND (("
    "Mid([Item No],6,1))=[FiberLetter]) AND (([Basic Product Construction].[Print Lin"
    "e 1])=\"optical fiber cable\") AND (([Basic Product Construction].Active)=True) "
    "AND (([Basic Product Construction].[Customer Part#]) Like \"*corning*\"))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblCorningPrint"
End
Begin OutputColumns
    Name ="[Basic Product Construction].[Print Line 2]"
    Expression ="[CorningPrint] & \" \" & Int(Mid([Item No],3,3)) & \" FIBER\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd520b5be5a557c4d8187e95a5b67aac9
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 1]"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 2]"
        dbInteger "ColumnWidth" ="5580"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =91
    Top =234
    Right =1842
    Bottom =708
    Left =-1
    Top =-1
    Right =1719
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =240
        Bottom =214
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =278
        Top =6
        Right =374
        Bottom =109
        Top =0
        Name ="tblCorningPrint"
        Name =""
    End
End
