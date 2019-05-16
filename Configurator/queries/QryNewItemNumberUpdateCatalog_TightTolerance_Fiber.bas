Operation =4
Option =8
Where ="(((tblItemCatalogInterface.ElementName)=\"QTS Print Vendor Data\") And ((tblFibe"
    "r_TightTolerance_Reference.TightToleranceFiber)=dbo_CableUnion.Fiber Or (tblFibe"
    "r_TightTolerance_Reference.TightToleranceFiber)=[Fiber2] Or (tblFiber_TightToler"
    "ance_Reference.TightToleranceFiber)=[Fiber3]))"
Begin InputTables
    Name ="tblFiber_TightTolerance_Reference"
    Name ="tblItemCatalogInterface"
    Name ="dbo_CableUnion"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="\"YES\""
End
Begin Joins
    LeftTable ="tblItemCatalogInterface"
    RightTable ="dbo_CableUnion"
    Expression ="tblItemCatalogInterface.ItemNumber = dbo_CableUnion.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x07606cd3abfd4948a3cd0616712c4d9f
End
Begin
    Begin
        dbText "Name" ="dbo_tblFiberExtensions.Fiber"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementName"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFiber_TightTolerance_Reference.TightToleranceFiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =107
    Top =103
    Right =1216
    Bottom =732
    Left =-1
    Top =-1
    Right =1077
    Bottom =389
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =90
        Top =81
        Right =310
        Bottom =305
        Top =0
        Name ="tblFiber_TightTolerance_Reference"
        Name =""
    End
    Begin
        Left =868
        Top =65
        Right =1276
        Bottom =362
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
    Begin
        Left =352
        Top =75
        Right =689
        Bottom =354
        Top =0
        Name ="dbo_CableUnion"
        Name =""
    End
End
