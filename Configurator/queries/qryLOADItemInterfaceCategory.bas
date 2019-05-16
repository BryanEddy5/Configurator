dbMemo "SQL" ="INSERT INTO tblItemInterfaceCategoryValues ( CategorySetName, CategoryValue, Des"
    "ignTypeID, Active )\015\012SELECT DISTINCT AFLPRD_INVItmCatg_CAB.CategorySetName"
    ", AFLPRD_INVItmCatg_CAB.CategoryName, 10 AS DesignTypeID, -1 AS Active\015\012FR"
    "OM AFLPRD_INVItmCatg_CAB LEFT JOIN tblItemInterfaceCategoryValues ON (AFLPRD_INV"
    "ItmCatg_CAB.CategoryName = tblItemInterfaceCategoryValues.CategoryValue) AND (AF"
    "LPRD_INVItmCatg_CAB.CategorySetName = tblItemInterfaceCategoryValues.CategorySet"
    "Name)\015\012WHERE (((AFLPRD_INVItmCatg_CAB.CategorySetName) Like \"new*\") AND "
    "((tblItemInterfaceCategoryValues.CategoryValue) Is Null) AND ((AFLPRD_INVItmCatg"
    "_CAB.ItemNumber) Like \"p*\"))\015\012ORDER BY AFLPRD_INVItmCatg_CAB.CategoryNam"
    "e;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xcd07954900be9f45b54df2da8dad1e20
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategoryName"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DesignTypeID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc62d964d5a1c1e46b4f5b2802cac9a99
        End
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1294bbd8402320438000a4a9ca25718b
        End
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategorySetName"
        dbLong "AggregateType" ="-1"
    End
End
