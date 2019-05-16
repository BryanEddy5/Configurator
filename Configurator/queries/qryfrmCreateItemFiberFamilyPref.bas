Operation =1
Option =0
Having ="(((qryFiberItemCustomer.Family)=Left([Forms]![frmCreateItem]![ItemNo],2)))"
Begin InputTables
    Name ="qryfrmCreateItemFiberType"
    Name ="qryFiberItemCustomer"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemFiberType.Fiber"
    Alias ="CountOfFiber"
    Expression ="Count(qryfrmCreateItemFiberType.Fiber)"
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
    Expression ="qryFiberItemCustomer.Family"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7b22ad4d892aa34d8f24c9548a73fd28
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfFiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c1e7d6fe6dbbc49a12258cb23786d1d
        End
    End
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6aaaaa630dc2c149ad3398e735f317cc
        End
    End
End
Begin
    State =0
    Left =57
    Top =279
    Right =1808
    Bottom =648
    Left =-1
    Top =-1
    Right =1719
    Bottom =138
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =343
        Top =6
        Right =628
        Bottom =154
        Top =0
        Name ="qryfrmCreateItemFiberType"
        Name =""
    End
    Begin
        Left =666
        Top =6
        Right =958
        Bottom =124
        Top =0
        Name ="qryFiberItemCustomer"
        Name =""
    End
End
