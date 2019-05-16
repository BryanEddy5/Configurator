dbMemo "SQL" ="SELECT tblAFLPRD_INVItmCatg_CAB.ItemNumber, tblAFLPRD_INVItmCatg_CAB.CategorySet"
    "Name, tblAFLPRD_INVItmCatg_CAB.CategoryName, tblAFLPRD_BOMInvComp_CAB.ComponentI"
    "temNumber, qryFiberOracle.Base, tblAFLPRD_BOMInvComp_CAB.AssemblyItemStatus\015\012"
    "FROM (tblAFLPRD_BOMInvComp_CAB INNER JOIN tblAFLPRD_INVItmCatg_CAB ON tblAFLPRD_"
    "BOMInvComp_CAB.AssemblyItemNumber=tblAFLPRD_INVItmCatg_CAB.ItemNumber) INNER JOI"
    "N qryFiberOracle ON tblAFLPRD_BOMInvComp_CAB.ComponentItemNumber=qryFiberOracle."
    "Oracle\015\012WHERE (((tblAFLPRD_INVItmCatg_CAB.ItemNumber) Like \"*ia\") AND (("
    "tblAFLPRD_INVItmCatg_CAB.CategorySetName)=\"new product type\") AND ((tblAFLPRD_"
    "BOMInvComp_CAB.ComponentItemNumber) Like \"p*\") AND ((tblAFLPRD_BOMInvComp_CAB."
    "AssemblyItemStatus)<>\"obsolete\" And (tblAFLPRD_BOMInvComp_CAB.AssemblyItemStat"
    "us) Not Like \"disc*\"))\015\012ORDER BY qryFiberOracle.Base;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8f5262d7541c274a8b38b2867588bb29
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVItmCatg_CAB.CategorySetName"
        dbInteger "ColumnWidth" ="4740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVItmCatg_CAB.CategoryName"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Base"
        dbInteger "ColumnWidth" ="3690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
