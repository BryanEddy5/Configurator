dbMemo "SQL" ="SELECT tblStandardsApplicationReference.Standard, tblCableConstructions.BaseID\015"
    "\012FROM tblCableConstructions INNER JOIN tblStandardsApplicationReference ON tb"
    "lCableConstructions.StandardsApplication = tblStandardsApplicationReference.Stan"
    "dardsApplication\015\012WHERE (((tblCableConstructions.BaseID)=667));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x76d7293d7c0ed24c8c61df4dcdce4de2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructions.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStandardsApplicationReference.Standard"
        dbLong "AggregateType" ="-1"
    End
End
