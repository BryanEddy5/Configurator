Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.BIF)=\"YES\") And ((qry_SearchByFiberType.[TBF/Ribbon])=Mid(Forms!frmCreateIte"
    "m!ItemNo,12,1)) And ((qry_SearchByFiberType.SpecialNotes)<>\"class 1\") And ((Fo"
    "rms!frmCreateItem!Fiber1Type)=qry_SearchByFiberType!Fiber_Type Or (Forms!frmCrea"
    "teItem!Fiber1Type)=qry_SearchByFiberType!Alt1_Fiber_Type Or (Forms!frmCreateItem"
    "!Fiber1Type)=qry_SearchByFiberType!Alt2_Fiber_Type Or (Forms!frmCreateItem!Fiber"
    "1Type)=qry_SearchByFiberType!Alt3_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)"
    "=qry_SearchByFiberType!Alt4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.SpecialNotes"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Alias ="Expr1"
    Expression ="Right(Left(Forms!frmCreateItem!ItemNo,6),1)"
    Expression ="qry_SearchByFiberType.SpecialNotes"
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
    0x7de4a0ce18ac34498be22425c7bde2b5
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
            0x5ddf9e2f6be7dd41b4a016bb78748eeb
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[OM#]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =199
    Right =1599
    Bottom =1104
    Left =-1
    Top =-1
    Right =1526
    Bottom =266
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
End
