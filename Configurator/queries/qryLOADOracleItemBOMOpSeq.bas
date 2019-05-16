dbMemo "SQL" ="INSERT INTO tblNetworkBOMOpSequencesInterface ( OrganizationCode, AssemblyItemNu"
    "mber, OperationSeqNum, AlternateRoutingDesignator, OperationCode, DepartmentCode"
    ", EffectivityDate, OperationDescription, StartupScrap, ShortText )\015\012SELECT"
    " AFLPRD_BOMOpSeq_CAB.OrganizationCode, AFLPRD_BOMOpSeq_CAB.AssemblyItemNumber, A"
    "FLPRD_BOMOpSeq_CAB.OperationSeqNum, AFLPRD_BOMOpSeq_CAB.AlternateRoutingDesignat"
    "or, AFLPRD_BOMOpSeq_CAB.OperationCode, AFLPRD_BOMOpSeq_CAB.DepartmentCode, AFLPR"
    "D_BOMOpSeq_CAB.EffectivityDate, AFLPRD_BOMOpSeq_CAB.OperationDescription, AFLPRD"
    "_BOMOpSeq_CAB.StartupScrap, AFLPRD_BOMOpSeq_CAB.ShortText\015\012FROM OraclePrem"
    "iseItems INNER JOIN AFLPRD_BOMOpSeq_CAB ON OraclePremiseItems.ItemNumber = AFLPR"
    "D_BOMOpSeq_CAB.AssemblyItemNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x47d492da1dc6004cbd38215eb38ec856
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
