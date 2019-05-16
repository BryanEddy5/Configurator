Operation =1
Option =0
Begin InputTables
    Name ="qryfrmCreateItemTBType"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemTBType.Family"
    Expression ="qryfrmCreateItemTBType.TBType"
    Expression ="qryfrmCreateItemTBType.JacketMaterial"
    Alias ="MaxOfCountOfTB Material"
    Expression ="Max(qryfrmCreateItemTBType.[CountOfTB Material])"
    Expression ="qryfrmCreateItemTBType.[TB Nominal OD]"
End
Begin Groups
    Expression ="qryfrmCreateItemTBType.Family"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType.TBType"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType.JacketMaterial"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType.[TB Nominal OD]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdd98eeeedc788f41aa9c22acc5bd0b53
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MaxOfCountOfTB Material"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e62b7f302882047a18aefe5e3578a1c
        End
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =56
    Top =313
    Right =1549
    Bottom =637
    Left =-1
    Top =-1
    Right =1461
    Bottom =59
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =243
        Bottom =122
        Top =0
        Name ="qryfrmCreateItemTBType"
        Name =""
    End
End
