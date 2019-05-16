dbMemo "SQL" ="SELECT DISTINCT tblDesignCodeFiberElementValue.FiberLetter, tblDesignCodeFiberEl"
    "ementValue.FiberValue, tblDesignCodeFiberElementValue.Mode, tblDesignCodeFiberEl"
    "ementValue.FiberExt, tblDesignCodeFiberElementValue.PreferedCableColor, tbl_Fibe"
    "rs.Active, tblDesignCodeFiberElementValue.Active\015\012FROM tblDesignCodeFiberE"
    "lementValue INNER JOIN (tbl_Fibers INNER JOIN qryFiberLetter ON tbl_Fibers.Fiber"
    " = qryFiberLetter.Fiber) ON tblDesignCodeFiberElementValue.FiberLetter = qryFibe"
    "rLetter.FiberLetter\015\012GROUP BY tblDesignCodeFiberElementValue.FiberLetter, "
    "tblDesignCodeFiberElementValue.FiberValue, tblDesignCodeFiberElementValue.Mode, "
    "tblDesignCodeFiberElementValue.FiberExt, tblDesignCodeFiberElementValue.Prefered"
    "CableColor, tbl_Fibers.Active, tblDesignCodeFiberElementValue.Active\015\012HAVI"
    "NG (((tblDesignCodeFiberElementValue.FiberLetter)<>\"@\") AND ((tbl_Fibers.Activ"
    "e)<>0) AND ((tblDesignCodeFiberElementValue.Active)<>0))\015\012ORDER BY tblDesi"
    "gnCodeFiberElementValue.FiberValue;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x18d3820d405805489118fae6ba10d095
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.FiberValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.Mode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.FiberLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.FiberExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.PreferedCableColor"
        dbLong "AggregateType" ="-1"
    End
End
