Operation =4
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Not Like \"*test*\") AND (("
    "[Basic Product Construction].[Print Line 1])=\"optical fiber cable\") AND (([Bas"
    "ic Product Construction].Active)=True) AND (([Basic Product Construction].[Custo"
    "mer Part#]) Not Like \"*corning*\") AND (([Basic Product Construction].[Fiber Or"
    "acle item])=[Fiber]) AND ((tbl_Fibers.[SM/MM]) Not Like \"*SM*\"))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Name ="[Basic Product Construction].[Print Line 2]"
    Expression ="[SM/MM] & \" \" & [OM#] & \" \" & IIf([BIF]=\"yes\",\"BIF \",\"\") & Int(Mid([It"
        "em No],3,3)) & \" FIBER\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x60631961a7c9664589a5021c3ee67eb7
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
    Bottom =880
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
        Left =412
        Top =6
        Right =802
        Bottom =254
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
