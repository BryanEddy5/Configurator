Operation =1
Option =0
Begin InputTables
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
    Alias ="Expr2"
    Expression ="[Basic (Cable/Sheath)].[Item No]"
    Alias ="Expr3"
    Expression ="[Basic (Cable/Sheath)].Base"
    Alias ="Expr4"
    Expression ="[Basic (Cable/Sheath)].[PSS Document #]"
    Alias ="Expr5"
    Expression ="[Basic (Cable/Sheath)].[Unit Series]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7515ea8c5600f84386c092a5b2d05dd5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Base"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9342d5111da1da4794258f4130710b78
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6e122ee8aa2984f82e4fc5f2261cfac
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5d2920eb6e66dd4084335ffc8f33e669
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xab1cb401e07de94db314e5bc0086da5f
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdf1e9ade62d6d94b90b186122d29024e
        End
    End
End
Begin
    State =0
    Left =76
    Top =229
    Right =1036
    Bottom =780
    Left =-1
    Top =-1
    Right =928
    Bottom =320
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =325
        Bottom =274
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End
