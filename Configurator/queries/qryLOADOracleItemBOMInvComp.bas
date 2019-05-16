dbMemo "SQL" ="INSERT INTO tblNetworkBOMInventoryCompsInterface ( OrganizationCode, AssemblyIte"
    "mNumber, ItemSeqNumber, OperationSeqNumber, ComponentItemNumber, EffectivityDate"
    ", ComponentQuantity, AlternateBOMDesignator, CountPerUOM, UnitID, LayerID )\015\012"
    "SELECT AFLPRD_BOMInvComp_CAB.OrganizationCode, AFLPRD_BOMInvComp_CAB.AssemblyIte"
    "mNumber, AFLPRD_BOMInvComp_CAB.ItemSeqNumber, AFLPRD_BOMInvComp_CAB.OperationSeq"
    "Number, AFLPRD_BOMInvComp_CAB.ComponentItemNumber, AFLPRD_BOMInvComp_CAB.Effecti"
    "vityDate, AFLPRD_BOMInvComp_CAB.ComponentQuantity, AFLPRD_BOMInvComp_CAB.Alterna"
    "teBOMDesignator, AFLPRD_BOMInvComp_CAB.CountPerUOM, AFLPRD_BOMInvComp_CAB.UnitId"
    ", AFLPRD_BOMInvComp_CAB.LayerId\015\012FROM OraclePremiseItems INNER JOIN AFLPRD"
    "_BOMInvComp_CAB ON OraclePremiseItems.ItemNumber = AFLPRD_BOMInvComp_CAB.Assembl"
    "yItemNumber\015\012ORDER BY AFLPRD_BOMInvComp_CAB.AssemblyItemNumber, AFLPRD_BOM"
    "InvComp_CAB.ItemSeqNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc41745319fdb2445a229ee217a29ef0d
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
        dbText "Name" ="AFLPRD_BOMInvComp_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMInvComp_CAB.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMInvComp_CAB.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
End
