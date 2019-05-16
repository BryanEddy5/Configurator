dbMemo "SQL" ="SELECT tblShipmentsCable.item_number, tblShipmentsCable.item_description, tblShi"
    "pmentsCable.bill_to_name, [Basic Product Construction].[Item No], [Basic Product"
    " Construction].Base, qryDesignCodeBreakdownAll.FiberType, qryDesignCodeBreakdown"
    "All.FiberType, tblShipmentsCable.transaction_date\015\012FROM qryDesignCodeBreak"
    "downAll INNER JOIN ([Basic Product Construction] INNER JOIN tblShipmentsCable ON"
    " [Basic Product Construction].[New Oracle Part #] = tblShipmentsCable.item_numbe"
    "r) ON qryDesignCodeBreakdownAll.Oracle = [Basic Product Construction].[New Oracl"
    "e Part #]\015\012WHERE (((tblShipmentsCable.bill_to_name) Like \"*fujikura*\") A"
    "ND (([Basic Product Construction].Base) Like \"sp024-3.8*\") AND ((qryDesignCode"
    "BreakdownAll.FiberType)=\"l\" Or (qryDesignCodeBreakdownAll.FiberType)=\"5\") AN"
    "D ((qryDesignCodeBreakdownAll.FiberType)=\"5\" Or (qryDesignCodeBreakdownAll.Fib"
    "erType)=\"l\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1332275bb24d184b9225c63f3581e909
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblShipmentsCable.item_description"
        dbInteger "ColumnWidth" ="10455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="6675"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
