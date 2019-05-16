dbMemo "SQL" ="INSERT INTO tblNetworkItemInterface ( OrganizationCode, ItemNumber, Description,"
    " PrimaryUOMCode, TemplateName, CatalogGroupName )\015\012SELECT AFLPRD_INVSysIte"
    "m_CAB.OrganizationCode, AFLPRD_INVSysItem_CAB.ItemNumber, AFLPRD_INVSysItem_CAB."
    "DESCRIPTION, AFLPRD_INVSysItem_CAB.PrimaryUOMCode, AFLPRD_INVSysItem_CAB.Templat"
    "eName, AFLPRD_INVSysItem_CAB.CatalogGroupName\015\012FROM AFLPRD_INVSysItem_CAB "
    "INNER JOIN OraclePremiseItems ON AFLPRD_INVSysItem_CAB.ItemNumber = OraclePremis"
    "eItems.ItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x20c1b23d99c26a459afa70fb662ea7bb
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
