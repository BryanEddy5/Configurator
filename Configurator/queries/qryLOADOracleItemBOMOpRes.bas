dbMemo "SQL" ="INSERT INTO tblNetworkBOMOpResourcesInterface ( OrganizationCode, AssemblyItemNu"
    "mber, AlternateRoutingDesignator, OperationSeqNumber, ResourceSeqNumber, Resourc"
    "eCode, EffectivityDate, StandardRateFlag, AssignedUnits, UsageRateOrAmount, Usag"
    "eRateOrAmountInverse, BasisType, ScheduleFlag, AutochargeType, ProcessFlag, Tran"
    "sactionType )\015\012SELECT AFLPRD_BOMOpRes_CAB.OrganizationCode, AFLPRD_BOMOpRe"
    "s_CAB.AssemblyItemNumber, AFLPRD_BOMOpRes_CAB.AlternateRoutingDesignator, AFLPRD"
    "_BOMOpRes_CAB.OperationSequenceNumber, AFLPRD_BOMOpRes_CAB.ResourceSeqNumber, AF"
    "LPRD_BOMOpRes_CAB.ResourceCode, AFLPRD_BOMOpRes_CAB.EffectivityDate, AFLPRD_BOMO"
    "pRes_CAB.StandardRateFlag, AFLPRD_BOMOpRes_CAB.AssignedUnits, AFLPRD_BOMOpRes_CA"
    "B.UsageRateOrAmount, AFLPRD_BOMOpRes_CAB.UsageRateOrAmountInverse, AFLPRD_BOMOpR"
    "es_CAB.BasisType, AFLPRD_BOMOpRes_CAB.ScheduleFlag, AFLPRD_BOMOpRes_CAB.AutoChar"
    "geType, AFLPRD_BOMOpRes_CAB.ProcessFlag, AFLPRD_BOMOpRes_CAB.TransactionType\015"
    "\012FROM AFLPRD_BOMOpRes_CAB INNER JOIN OraclePremiseItems ON AFLPRD_BOMOpRes_CA"
    "B.AssemblyItemNumber = OraclePremiseItems.ItemNumber\015\012WHERE (((AFLPRD_BOMO"
    "pRes_CAB.ResourceSeqNumber) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x3cb8a1d1ad0eca42933289a3ad5d9f08
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
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
