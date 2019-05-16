dbMemo "SQL" ="SELECT Left(tblAFLPRD_INVItmCtlg_CAB!Item,2) AS Prefix, tblAFLPRD_INVItmCtlg_CAB"
    ".ElementName, tblAFLPRD_INVItmCtlg_CAB.ElementValue, Count(tblAFLPRD_INVItmCtlg_"
    "CAB.ElementValue) AS CountOfElementValue, qryFiberOracle.Base\015\012FROM qryFib"
    "erOracle INNER JOIN tblAFLPRD_INVItmCtlg_CAB ON qryFiberOracle.Oracle=tblAFLPRD_"
    "INVItmCtlg_CAB.Item\015\012GROUP BY Left(tblAFLPRD_INVItmCtlg_CAB!Item,2), tblAF"
    "LPRD_INVItmCtlg_CAB.ElementName, tblAFLPRD_INVItmCtlg_CAB.ElementValue, qryFiber"
    "Oracle.Base\015\012HAVING (((tblAFLPRD_INVItmCtlg_CAB.ElementName)=\"cable type\""
    ") And ((Count(tblAFLPRD_INVItmCtlg_CAB.ElementValue))>10) And ((Left(tblAFLPRD_I"
    "NVItmCtlg_CAB!Item,2)) Like \"p*\"))\015\012ORDER BY Left(tblAFLPRD_INVItmCtlg_C"
    "AB!Item,2), Count(tblAFLPRD_INVItmCtlg_CAB.ElementValue) DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x71b71986773bb24f823617b666e250bf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVItmCtlg_CAB.ElementValue"
        dbInteger "ColumnWidth" ="3465"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Prefix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6919485c5e26ba4bb6307eaf05154721
        End
    End
    Begin
        dbText "Name" ="CountOfElementValue"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcb49fa7f79b3d8499c11a80e4e3bf5fc
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5405d3cc941845428a66f7794631862f
        End
    End
End
