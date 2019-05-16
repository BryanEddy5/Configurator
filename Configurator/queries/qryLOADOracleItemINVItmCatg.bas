dbMemo "SQL" ="INSERT INTO tblNetworkItemCategoryInterface ( OrganizationCode, ItemNumber, Cate"
    "gorySetName, CategoryName, AllowOverride )\015\012SELECT AFLPRD_INVItmCatg_CAB.O"
    "rganizationCode, AFLPRD_INVItmCatg_CAB.ItemNumber, AFLPRD_INVItmCatg_CAB.Categor"
    "ySetName, AFLPRD_INVItmCatg_CAB.CategoryName, AFLPRD_INVItmCatg_CAB.AllowOverrid"
    "e\015\012FROM AFLPRD_INVItmCatg_CAB INNER JOIN OraclePremiseItems ON AFLPRD_INVI"
    "tmCatg_CAB.ItemNumber = OraclePremiseItems.ItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc008bf5d7ed5c44a8586b554e50a46a4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="AFLPRD_BOMInvComp_CAB.CompItemStatus"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategorySetName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategoryName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.AllowOverride"
        dbLong "AggregateType" ="-1"
    End
End
