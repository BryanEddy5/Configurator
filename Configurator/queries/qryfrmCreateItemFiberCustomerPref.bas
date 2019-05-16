Operation =1
Option =0
Having ="(((qryFiberItemCustomer.Customer)=[Forms]![frmCreateItem]![Customer]))"
Begin InputTables
    Name ="qryfrmCreateItemFiberType"
    Name ="qryFiberItemCustomer"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemFiberType.Fiber"
    Alias ="CountOfFiber"
    Expression ="Count(qryfrmCreateItemFiberType.Fiber)"
    Expression ="qryFiberItemCustomer.Customer"
End
Begin Joins
    LeftTable ="qryfrmCreateItemFiberType"
    RightTable ="qryFiberItemCustomer"
    Expression ="qryfrmCreateItemFiberType.Fiber = qryFiberItemCustomer.Fiber"
    Flag =1
End
Begin OrderBy
    Expression ="Count(qryfrmCreateItemFiberType.Fiber)"
    Flag =1
End
Begin Groups
    Expression ="qryfrmCreateItemFiberType.Fiber"
    GroupLevel =0
    Expression ="qryFiberItemCustomer.Customer"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x24e4f4c21b4b194a9fd74f6aad231201
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfFiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7598f4c5085e4b4cb54c5ac01fdf7779
        End
    End
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x08e7df5accd4624a8460534cc8acb96d
        End
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf8ca62919f714b4699bad7b2ec889da1
        End
    End
End
Begin
    State =0
    Left =111
    Top =398
    Right =1723
    Bottom =829
    Left =-1
    Top =-1
    Right =1580
    Bottom =200
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =567
        Bottom =154
        Top =0
        Name ="qryfrmCreateItemFiberType"
        Name =""
    End
    Begin
        Left =669
        Top =16
        Right =906
        Bottom =134
        Top =0
        Name ="qryFiberItemCustomer"
        Name =""
    End
End
