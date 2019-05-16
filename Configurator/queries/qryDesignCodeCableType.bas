dbMemo "SQL" ="SELECT DISTINCT tblDesignCodeCableType.CableLetter, tblDesignCodeCableType.Cable"
    "Type, tblDesignCodeCableType.CablePasses\015\012FROM qryDesignCodeBreakdown INNE"
    "R JOIN tblDesignCodeCableType ON qryDesignCodeBreakdown.CableType=tblDesignCodeC"
    "ableType.CableLetter\015\012WHERE (((tblDesignCodeCableType.CableType) Not Like "
    "\"*center*\"))\015\012ORDER BY tblDesignCodeCableType.CableType;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x40f970a71e59764db6c3c748268c076a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
