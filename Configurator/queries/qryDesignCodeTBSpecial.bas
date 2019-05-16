dbMemo "SQL" ="SELECT tblTightBuffer.TBLetter, tblTightBuffer.TBDesc\015\012FROM tblTightBuffer"
    " LEFT JOIN qryDesignCodeTB ON tblTightBuffer.TBLetter = qryDesignCodeTB.TBLetter"
    "\015\012WHERE (((qryDesignCodeTB.TBLetter) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x07a16ea923dee149966c0ff12c5fcf34
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblTightBuffer.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTightBuffer.TBDesc"
        dbLong "AggregateType" ="-1"
    End
End
