dbMemo "SQL" ="INSERT INTO tblNetworkBOMRoutingsInterface ( OrganizationCode, AssemblyItemNumbe"
    "r, CompletionSubinventory, CommonItemNumber, AlternateRoutingDesignator, ItemCop"
    "ied, LocationName )\015\012SELECT AFLPRD_BOMRoutSub_CAB.OrganizationCode, AFLPRD"
    "_BOMRoutSub_CAB.AssemblyItemNumber, AFLPRD_BOMRoutSub_CAB.CompletionSubinventory"
    ", AFLPRD_BOMRoutSub_CAB.CommonItemNumber, AFLPRD_BOMRoutSub_CAB.AlternateRouting"
    "Designator, AFLPRD_BOMRoutSub_CAB.ItemCopied, AFLPRD_BOMRoutSub_CAB.LocationName"
    "\015\012FROM OraclePremiseItems INNER JOIN AFLPRD_BOMRoutSub_CAB ON OraclePremis"
    "eItems.ItemNumber = AFLPRD_BOMRoutSub_CAB.AssemblyItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5310755d4ac1784fa7e5ff73c6f1ee99
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
        dbText "Name" ="AFLPRD_BOMRoutSub_CAB.Number"
        dbLong "AggregateType" ="-1"
    End
End
