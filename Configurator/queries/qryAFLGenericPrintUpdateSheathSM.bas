Operation =1
Option =0
Where ="((([Basic (Cable/Sheath)].[New Oracle Part Number]) Not Like \"*test*\") And ((["
    "Basic (Cable/Sheath)].[Print Line 1])=\"optical fiber cable\") And (([Basic (Cab"
    "le/Sheath)].[Customer Part Num]) Not Like \"*CORNING*\") And ((Left([Fiber Type "
    "A],8))=[Fiber]) And ((tbl_Fibers.[SM/MM]) Not Like \"*MM*\") And (([Basic (Cable"
    "/Sheath)].[Item No]) Not Like \"*/*\") And (([Basic (Cable/Sheath)].Active)=-1))"
Begin InputTables
    Name ="tbl_Fibers"
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="[Basic (Cable/Sheath)].[Print Line 2]"
    Alias ="Expr2"
    Expression ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
    Alias ="Expr1"
    Expression ="Mid([Item No],6,1)"
    Alias ="Expr3"
    Expression ="[Basic (Cable/Sheath)].[Print Line 1]"
    Alias ="Expr4"
    Expression ="[Basic (Cable/Sheath)].Active"
    Alias ="Expr5"
    Expression ="[Basic (Cable/Sheath)].Customer"
    Alias ="Expr6"
    Expression ="[Basic (Cable/Sheath)].[Customer Part Num]"
    Alias ="Expr2"
    Expression ="Left([Fiber Type A],8)"
    Expression ="tbl_Fibers.[SM/MM]"
    Alias ="Expr7"
    Expression ="[Basic (Cable/Sheath)].[Item No]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x5da14633852cc94182674afcc31ad354
End
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="qryAFLGenericPrintUpdateSheathSM.Expr1"
dbBoolean "OrderByOn" ="-1"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x98667b373c59e347b0ca5d17f608b787
        End
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 2]"
        dbInteger "ColumnWidth" ="5580"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 1]"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Active"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Customer"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Customer Part Num]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4598abe4fd424c429a1d17af2599f04a
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6552b9f30febba448fa87b5de62b4b37
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x34483bfa1e69444bbc2b1289aef433f8
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x63a7d91ef041e34d8af08c63d675f74c
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8f033f110898049a52c887e4d6ae9ea
        End
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdeb8f473a630764fa062526e34747ae8
        End
    End
End
Begin
    State =0
    Left =53
    Top =192
    Right =1905
    Bottom =900
    Left =-1
    Top =-1
    Right =1820
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =575
        Top =5
        Right =965
        Bottom =253
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =278
        Top =6
        Right =374
        Bottom =124
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End
