Operation =4
Option =0
Where ="(((tbl_Fibers.Fiber)=[Forms]![frmCreateItem]![Fiber1]))"
Begin InputTables
    Name ="tbl_Fibers"
    Name ="tblCreateItem"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="tblCreateItem.BIF"
    Expression ="tbl_Fibers!BIF"
    Alias ="Expr2"
    Name ="tblCreateItem.OM"
    Expression ="[OM#]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x42789079901f1148acff622d520548e3
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c28e51dc1e7c84b8ea2df733f40c387
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x13580706c44f044894b9d9a50578040b
        End
    End
End
Begin
    State =0
    Left =114
    Top =278
    Right =1822
    Bottom =746
    Left =-1
    Top =-1
    Right =1676
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =371
        Bottom =214
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =453
        Top =12
        Right =746
        Bottom =55
        Top =0
        Name ="tblCreateItem"
        Name =""
    End
End
