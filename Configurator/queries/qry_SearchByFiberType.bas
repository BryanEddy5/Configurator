Operation =1
Option =2
Where ="(((tbl_Fibers.Active)<>0))"
Begin InputTables
    Name ="tbl_Fibers"
    Name ="tblFiberTBCatalog"
    Name ="tblFiberTBIndicators"
End
Begin OutputColumns
    Expression ="tbl_Fibers.Fiber"
    Expression ="tbl_Fibers.[SM/MM]"
    Expression ="tbl_Fibers.BIF"
    Expression ="tbl_Fibers.[OM#]"
    Expression ="tbl_Fibers.Fiber_Type"
    Expression ="tbl_Fibers.Alt1_Fiber_Type"
    Expression ="tbl_Fibers.Alt2_Fiber_Type"
    Expression ="tbl_Fibers.Alt3_Fiber_Type"
    Expression ="tbl_Fibers.Description_Oracle"
    Expression ="tbl_Fibers.Alt4_Fiber_Type"
    Expression ="tbl_Fibers.Cost"
    Expression ="tbl_Fibers.Vendor"
    Expression ="tbl_Fibers.ITU"
    Expression ="tbl_Fibers.SpecialNotes"
    Expression ="tbl_Fibers.Extension"
    Expression ="tbl_Fibers.Active"
    Alias ="TBF/Ribbon"
    Expression ="tblFiberTBIndicators.TBLetter"
    Expression ="tbl_Fibers.TBCatalogType"
End
Begin Joins
    LeftTable ="tbl_Fibers"
    RightTable ="tblFiberTBCatalog"
    Expression ="tbl_Fibers.TBCatalogType = tblFiberTBCatalog.TBCatalogType"
    Flag =1
    LeftTable ="tblFiberTBIndicators"
    RightTable ="tblFiberTBCatalog"
    Expression ="tblFiberTBIndicators.TBCatalogType = tblFiberTBCatalog.TBCatalogType"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb50eda4766fb2d4dbfffe3a86792b54a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([qry_SearchByFiberType].[Fiber] Like \"*286*\")"
Begin
    Begin
        dbText "Name" ="TBF/Ribbon"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4945575080422d438cd1a67fb3d0dd31
        End
    End
    Begin
        dbText "Name" ="tbl_Fibers.[SM/MM]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.BIF"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.[OM#]"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Fiber_Type"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Alt1_Fiber_Type"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Alt2_Fiber_Type"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Alt3_Fiber_Type"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Description_Oracle"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Alt4_Fiber_Type"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Fiber"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Cost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Extension"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Vendor"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="tbl_Fibers.ITU"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="tbl_Fibers.SpecialNotes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Active"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="tbl_Fibers.TBCatalogType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbe8d7e4d7e22344a9184b4748b2d3818
        End
    End
    Begin
        dbText "Name" ="SM/MM"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4696852d25a25040887e2fb2406c6e38
        End
    End
    Begin
        dbText "Name" ="BIF"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1109860b25703a44adb9a10e77b24e3a
        End
    End
    Begin
        dbText "Name" ="OM#"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5e12a51bc380db4ba14e480f0f64e3cb
        End
    End
    Begin
        dbText "Name" ="Fiber_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9b8286f4145e69448b8991061c73ce6c
        End
    End
    Begin
        dbText "Name" ="Alt1_Fiber_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x38e9b01404ef184ab1a3101cfd6323cc
        End
    End
    Begin
        dbText "Name" ="Alt2_Fiber_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f7b810f67a7b4459f5254c5780864c7
        End
    End
    Begin
        dbText "Name" ="Alt3_Fiber_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x510f518451707e48a056499dc8a532a3
        End
    End
    Begin
        dbText "Name" ="Description_Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52def675f6ea5d48b53a72aa3fa13a31
        End
    End
    Begin
        dbText "Name" ="Alt4_Fiber_Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9de3d7a4e8c4164e920d61b29ac0cd43
        End
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3262bffc5c6182458312a22493c0a27c
        End
    End
    Begin
        dbText "Name" ="Vendor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd3a4d892752169458d343a5f3d11397f
        End
    End
    Begin
        dbText "Name" ="ITU"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80e0e24497068c448302b64348cf81be
        End
    End
    Begin
        dbText "Name" ="SpecialNotes"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xccdcd6d303202147b5a239a7793e9828
        End
    End
    Begin
        dbText "Name" ="Extension"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x60b35315b71d0b4d9cbd6312854c88ad
        End
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x292f8b9eec035640ad6f876a979d9983
        End
    End
    Begin
        dbText "Name" ="TBCatalogType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb42be2702c7ef048be29bed047abc7ee
        End
    End
End
Begin
    State =0
    Left =48
    Top =244
    Right =1565
    Bottom =998
    Left =-1
    Top =-1
    Right =1485
    Bottom =116
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =231
        Top =3
        Right =408
        Bottom =157
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =534
        Top =19
        Right =678
        Bottom =163
        Top =0
        Name ="tblFiberTBCatalog"
        Name =""
    End
    Begin
        Left =851
        Top =23
        Right =995
        Bottom =164
        Top =0
        Name ="tblFiberTBIndicators"
        Name =""
    End
End
