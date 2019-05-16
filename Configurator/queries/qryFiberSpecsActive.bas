dbMemo "SQL" ="SELECT tblAFLPRD_INVSysItemSpec_CAB.*, tblAFLPRD_INVSysItemCost_CAB.Item_Status "
    "AS FiberStatus\015\012FROM tblAFLPRD_INVSysItemSpec_CAB INNER JOIN tblAFLPRD_INV"
    "SysItemCost_CAB ON tblAFLPRD_INVSysItemSpec_CAB.FiberType = tblAFLPRD_INVSysItem"
    "Cost_CAB.ItemNumber\015\012WHERE (((tblAFLPRD_INVSysItemSpec_CAB.SpecificationEl"
    "ement) Like \"*attn*\") AND ((tblAFLPRD_INVSysItemCost_CAB.Item_Status)<>\"obsol"
    "ete\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x60af27b3a7b76f4d9074285378f785bc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="4260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FiberStatus"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7fee6144505fef4fb1b1a5d9d8dc9faa
        End
    End
End
