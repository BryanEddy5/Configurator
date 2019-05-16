dbMemo "SQL" ="INSERT INTO tblNetworkItemCatalogInterface ( ItemNumber, ElementName, ElementVal"
    "ue, AllowOverride )\015\012SELECT AFLPRD_INVItmCtlg_CAB.Item, AFLPRD_INVItmCtlg_"
    "CAB.ElementName, AFLPRD_INVItmCtlg_CAB.ElementValue, AFLPRD_INVItmCtlg_CAB.Allow"
    "Override\015\012FROM OraclePremiseItems INNER JOIN AFLPRD_INVItmCtlg_CAB ON Orac"
    "lePremiseItems.ItemNumber = AFLPRD_INVItmCtlg_CAB.Item;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf6ace1d8ceae9f459c50986b4e394432
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
End
