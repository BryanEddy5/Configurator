dbMemo "SQL" ="INSERT INTO tblNetworkBOMBillOfMtlsInterface ( ItemNumber, OrganizationCode, Ass"
    "emblyType, AlternateBOMDesignator, CommonOrgCode, CommonItemNumber )\015\012SELE"
    "CT AFLPRD_BOMBillMtl_CAB.Assembly, AFLPRD_BOMBillMtl_CAB.OrganizationCode, AFLPR"
    "D_BOMBillMtl_CAB.AssemblyType, AFLPRD_BOMBillMtl_CAB.AlternateBOMDesignator, AFL"
    "PRD_BOMBillMtl_CAB.CommonOrgCode, AFLPRD_BOMBillMtl_CAB.CommonBOM\015\012FROM AF"
    "LPRD_BOMBillMtl_CAB INNER JOIN OraclePremiseItems ON AFLPRD_BOMBillMtl_CAB.Assem"
    "bly = OraclePremiseItems.ItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf5aba635e951924e8ecc696094d03216
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
