dbMemo "SQL" ="SELECT tblAFLPRD_INVSysItemSpec_CAB.ItemNumber, tblAFLPRD_INVSysItemSpec_CAB.Spe"
    "cName, tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement, tblAFLPRD_INVSysItemSp"
    "ec_CAB.UOM, \"(\" & qryFiberOracle.ListingCompany & \")c(\" & qryFiberOracle.Lis"
    "tingCompany & \")\" AS TargetValue, tblAFLPRD_INVSysItemSpec_CAB.LowerLimit, tbl"
    "AFLPRD_INVSysItemSpec_CAB.UpperLimit, tblAFLPRD_INVSysItemSpec_CAB.UserLowerLimi"
    "t, tblAFLPRD_INVSysItemSpec_CAB.UserUpperLimit, tblAFLPRD_INVSysItemSpec_CAB.Fib"
    "erType, tblAFLPRD_INVSysItemSpec_CAB.TestLength, tblAFLPRD_INVSysItemSpec_CAB.Te"
    "stIncrement\015\012FROM qryFiberOracle INNER JOIN tblAFLPRD_INVSysItemSpec_CAB O"
    "N qryFiberOracle.Oracle = tblAFLPRD_INVSysItemSpec_CAB.ItemNumber\015\012WHERE ("
    "((tblAFLPRD_INVSysItemSpec_CAB.SpecName) Not Like \"*-00\") AND ((tblAFLPRD_INVS"
    "ysItemSpec_CAB.SpecificationElement) Like \"*label*\") AND ((tblAFLPRD_INVSysIte"
    "mSpec_CAB.TargetValue) Not Like \"*\" & [ListingCompany] & \"*\") AND ((qryFiber"
    "Oracle.ListingCompany) Like \"*UL*\" Or (qryFiberOracle.ListingCompany) Like \"*"
    "etl*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb4e884634db33144acc9f6a9ce915c01
End
dbMemo "OrderBy" ="qrySpecListingComparison_Oracle.TargetValue"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetValue"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9777f2467501d54ea4851ca1fde97f06
        End
    End
End
