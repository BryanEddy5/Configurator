dbMemo "SQL" ="INSERT INTO tblNetworkBOMRoutingsAttachments ( OrganizationCode, AssemblyItemNum"
    "ber, OperationSeqNumber, SeqNumber, ShortText, Description )\015\012SELECT AFLPR"
    "D_BOMRoutAtc_CAB.OrganizationCode, AFLPRD_BOMRoutAtc_CAB.AssemblyItemNumber, AFL"
    "PRD_BOMRoutAtc_CAB.OperationSeqNumber, AFLPRD_BOMRoutAtc_CAB.SeqNumber, AFLPRD_B"
    "OMRoutAtc_CAB.ShortText, AFLPRD_BOMRoutAtc_CAB.Description\015\012FROM AFLPRD_BO"
    "MRoutAtc_CAB INNER JOIN OraclePremiseItems ON AFLPRD_BOMRoutAtc_CAB.AssemblyItem"
    "Number = OraclePremiseItems.ItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4874ea468ac0474bbf00e3031c46dc14
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
