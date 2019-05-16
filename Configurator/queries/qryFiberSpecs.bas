dbMemo "SQL" ="SELECT tblAFLPRD_INVSysItemSpec_CAB.FiberType, Count(tblAFLPRD_INVSysItemSpec_CA"
    "B.FiberType) AS FiberTypeCount, tblAFLPRD_INVSysItemSpec_CAB.SpecificationElemen"
    "t, tblAFLPRD_INVSysItemSpec_CAB.LowerLimit, tblAFLPRD_INVSysItemSpec_CAB.UpperLi"
    "mit, tblAFLPRD_INVSysItemSpec_CAB.UserLowerLimit, tblAFLPRD_INVSysItemSpec_CAB.U"
    "serUpperLimit, tblAFLPRD_INVSysItemCost_CAB.Item_Status INTO tblFiberSpecs\015\012"
    "FROM tblAFLPRD_INVSysItemSpec_CAB INNER JOIN tblAFLPRD_INVSysItemCost_CAB ON tbl"
    "AFLPRD_INVSysItemSpec_CAB.FiberType=tblAFLPRD_INVSysItemCost_CAB.ItemNumber\015\012"
    "GROUP BY tblAFLPRD_INVSysItemSpec_CAB.FiberType, tblAFLPRD_INVSysItemSpec_CAB.Sp"
    "ecificationElement, tblAFLPRD_INVSysItemSpec_CAB.LowerLimit, tblAFLPRD_INVSysIte"
    "mSpec_CAB.UpperLimit, tblAFLPRD_INVSysItemSpec_CAB.UserLowerLimit, tblAFLPRD_INV"
    "SysItemSpec_CAB.UserUpperLimit, tblAFLPRD_INVSysItemCost_CAB.Item_Status\015\012"
    "HAVING (((tblAFLPRD_INVSysItemSpec_CAB.FiberType) Like \"FBR*\") AND ((tblAFLPRD"
    "_INVSysItemSpec_CAB.SpecificationElement) Like \"*ATTN*\" And (tblAFLPRD_INVSysI"
    "temSpec_CAB.SpecificationElement) Not Like \"*ALT*\" And (tblAFLPRD_INVSysItemSp"
    "ec_CAB.SpecificationElement) Like \"*FIBER1*\") AND ((tblAFLPRD_INVSysItemCost_C"
    "AB.Item_Status)=\"ACTIVE\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf4efa0a7b6f78944affa8ea6950fb950
End
dbBoolean "UseTransaction" ="-1"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="4080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FiberTypeCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x49e948cbd470894885f4c5f002899143
        End
    End
End
