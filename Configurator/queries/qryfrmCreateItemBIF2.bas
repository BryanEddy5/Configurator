Operation =1
Option =0
Where ="(((tbl_Fibers.Fiber)=[Forms]![frmCreateItem]![Fiber2]))"
Begin InputTables
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tbl_Fibers.BIF"
    Alias ="Expr2"
    Expression ="tbl_Fibers.Fiber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0dda285c7d6fb1488e07beebfc8b4fa6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd2fa2bec6af66f489845af4b02a55a6e
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf55d8f78b1dfdf40be9c9faff30d456e
        End
    End
End
Begin
    State =0
    Left =199
    Top =156
    Right =1898
    Bottom =480
    Left =-1
    Top =-1
    Right =1667
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
