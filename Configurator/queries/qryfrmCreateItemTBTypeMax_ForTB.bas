Operation =1
Option =0
Begin InputTables
    Name ="qryfrmCreateItemTBType_ForTB"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemTBType_ForTB.Family"
    Expression ="qryfrmCreateItemTBType_ForTB.TBType"
    Expression ="qryfrmCreateItemTBType_ForTB.[Jacket Material]"
    Alias ="MaxOfCountOfTB Material"
    Expression ="Max(qryfrmCreateItemTBType_ForTB.[CountOfTB Material])"
    Expression ="qryfrmCreateItemTBType_ForTB.[TB Nominal OD]"
End
Begin Groups
    Expression ="qryfrmCreateItemTBType_ForTB.Family"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_ForTB.TBType"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_ForTB.[Jacket Material]"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_ForTB.[TB Nominal OD]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9b3ba98e1771a44abf8262c4f325abb1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdee507653e1e75439d40bcd64143e498
        End
    End
    Begin
        dbText "Name" ="TBType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe93daaff886fb0429c9c8a117175cd29
        End
    End
    Begin
        dbText "Name" ="MaxOfCountOfTB Material"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2863bc5ae6388b42aadda802147f8a13
        End
    End
End
Begin
    State =0
    Left =278
    Top =222
    Right =1771
    Bottom =546
    Left =-1
    Top =-1
    Right =1461
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =281
        Top =6
        Right =377
        Bottom =124
        Top =0
        Name ="qryfrmCreateItemTBType_ForTB"
        Name =""
    End
End
