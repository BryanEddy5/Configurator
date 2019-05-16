Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.[TBF/Ribbon])=Mid(Forms!frmCreateItem!ItemNo,12,1)) And ((Forms!frmCreateItem!"
    "Fiber1Type)=qry_SearchByFiberType!Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)"
    "=qry_SearchByFiberType!Alt1_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_S"
    "earchByFiberType!Alt2_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchB"
    "yFiberType!Alt3_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFiber"
    "Type!Alt4_Fiber_Type) And ((Mid(Forms!frmCreateItem!ItemNo,14,30)) Like \"*\" & "
    "RTrim([FiberExtension]) & \"-*\" Or (Mid(Forms!frmCreateItem!ItemNo,14,30)) Like"
    " \"*\" & RTrim([FiberExtension])))"
Begin InputTables
    Name ="qry_SearchByFiberType"
    Name ="tblFiberExtensions"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="tblFiberExtensions.FiberExtension"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.SpecialNotes"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Alias ="Expr1"
    Expression ="Right(Left(Forms!frmCreateItem!ItemNo,6),1)"
    Alias ="ItemExt"
    Expression ="Mid(Forms!frmCreateItem!ItemNo,14,30)"
End
Begin Joins
    LeftTable ="qry_SearchByFiberType"
    RightTable ="tblFiberExtensions"
    Expression ="qry_SearchByFiberType.Fiber = tblFiberExtensions.Fiber"
    Flag =1
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
    0x3485e43b6cf2e64cb871fb8c73d9c5ce
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
            0x13a9cb80d5937245bd638954d323eede
        End
    End
    Begin
        dbText "Name" ="ItemExt"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa2cb706154ce6b4f8fa28d084cf1963a
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[OM#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.BIF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Description_Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Cost"
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
        dbText "Name" ="tblFiberExtensions.FiberExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFiberExtensions.[FiberExtension]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1551
    Bottom =957
    Left =-1
    Top =-1
    Right =1519
    Bottom =336
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =331
        Bottom =244
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
    Begin
        Left =607
        Top =9
        Right =751
        Bottom =153
        Top =0
        Name ="tblFiberExtensions"
        Name =""
    End
End
