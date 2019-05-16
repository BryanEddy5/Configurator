Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.Description_Oracle) Not Like \"*class 1*\") And ((qry_SearchByFiberType.[TBF/R"
    "ibbon])=Mid(Forms!frmCreateItem!ItemNo,12,1)) And ((Forms!frmCreateItem!Fiber1Ty"
    "pe)=qry_SearchByFiberType!Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_Sea"
    "rchByFiberType!Alt1_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByF"
    "iberType!Alt2_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberTy"
    "pe!Alt3_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Alt"
    "4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.Extension"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.SpecialNotes"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.[OM#]"
    Alias ="Expr1"
    Expression ="Right(Left(Forms!frmCreateItem!ItemNo,6),1)"
End
Begin OrderBy
    Expression ="qry_SearchByFiberType.Cost"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2e4ff37aeb025a45b77dbd9b46c6815b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qryfrmCreateItemFiberType].[Fiber]"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x29044e9aa449864baef8485bd1b186ab
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.SpecialNotes"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc99eaab4c0670241ad32f07ad8066d3f
        End
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d8f03a2faf8e341bbcafd31bb276043
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.BIF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Description_Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[OM#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Extension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[TBF/Ribbon]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[SM/MM]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =53
    Top =146
    Right =1588
    Bottom =796
    Left =-1
    Top =-1
    Right =1503
    Bottom =350
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =411
        Bottom =304
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
End
