Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.[TBF/Ribbon])=[TB/Ribbon_Type]) And ((tbl_FindFiber_byLetters.Fiber_Letter)=[F"
    "iber_Type] Or (tbl_FindFiber_byLetters.Fiber_Letter)=[Alt1_Fiber_Type] Or (tbl_F"
    "indFiber_byLetters.Fiber_Letter)=[Alt2_Fiber_Type] Or (tbl_FindFiber_byLetters.F"
    "iber_Letter)=[Alt3_Fiber_Type]) And ((tbl_FindFiber_byLetters.Active)=-1))"
Begin InputTables
    Name ="qry_SearchByFiberType"
    Name ="tbl_FindFiber_byLetters"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Alias ="Expr1"
    Expression ="tbl_FindFiber_byLetters.Fiber_Letter"
    Alias ="TB"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.BIF"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb3e1a452d8bd3548a3283ff9b12217e4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TB"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc5d862ef38e14f4b8bea2892e9103ab0
        End
    End
    Begin
        dbText "Name" ="tbl_FindFiber_byLetters.Fiber_Letter"
        dbInteger "ColumnWidth" ="1425"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e56178693ba174aab5ed63af5006c9c
        End
    End
End
Begin
    State =0
    Left =124
    Top =213
    Right =1705
    Bottom =757
    Left =-1
    Top =-1
    Right =1549
    Bottom =296
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =584
        Top =21
        Right =921
        Bottom =214
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
    Begin
        Left =71
        Top =22
        Right =432
        Bottom =125
        Top =0
        Name ="tbl_FindFiber_byLetters"
        Name =""
    End
End
