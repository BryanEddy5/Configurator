dbMemo "SQL" ="SELECT [New Oracle Part #] AS OracleItemNumber, [Item No] AS DesignCode, Custome"
    "r, Active, Base\015\012FROM [Basic Product Construction];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x240e8e48892fea47b6dff8fad86a77a8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="OracleItemNumber"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf70b5c553fddf541b03f2b6ecf6d13ef
        End
    End
    Begin
        dbText "Name" ="Customer"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8176d5bf6d67104fba641cc813fe893a
        End
    End
    Begin
        dbText "Name" ="Active"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8f01d671a617aa4cbf6abae03ce62d02
        End
    End
    Begin
        dbText "Name" ="DesignCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x64d05cbea304d6498eb721562170af69
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x32d39f1cfa39d343b0015790fdea5368
        End
    End
End
