dbMemo "SQL" ="UPDATE AFLPRD_INVSysItemCost_CAB INNER JOIN [Basic (Cable/Sheath)] ON AFLPRD_INV"
    "SysItemCost_CAB.ItemNumber = [Basic (Cable/Sheath)].[New Oracle Part Number] SET"
    " [Basic (Cable/Sheath)].OracleStatus = [Item_Status];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x2c600fb04a8f3745a9c3f6552447bf26
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].OracleStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
End
