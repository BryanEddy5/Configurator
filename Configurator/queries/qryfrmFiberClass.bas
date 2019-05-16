Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.SpecialNotes)=\"CLASS 1\") And ((qry_SearchByFiberType.[TBF/Ribbon])=Mid(Forms"
    "!frmCreateItem!ItemNo,12,1)) And ((Forms!frmCreateItem!Fiber1Type)=qry_SearchByF"
    "iberType!Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Al"
    "t1_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Alt2_Fib"
    "er_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Alt3_Fiber_Typ"
    "e Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Alt4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.Extension"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.SpecialNotes"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.[OM#]"
    Alias ="Expr1"
    Expression ="Right(Left(Forms!frmCreateItem!ItemNo,6),1)"
    Expression ="qry_SearchByFiberType.ITU"
End
Begin OrderBy
    Expression ="qry_SearchByFiberType.Cost"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcfafe2b655bb9d4dab4b8fe01f754d52
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_SearchByFiberType.SpecialNotes"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x96529d3c10ef6f42823769d73f98d138
        End
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
        dbText "Name" ="qry_SearchByFiberType.[TBF/Ribbon]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[SM/MM]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.ITU"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =10
    Top =152
    Right =1628
    Bottom =683
    Left =-1
    Top =-1
    Right =1586
    Bottom =283
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
