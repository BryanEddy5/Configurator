dbMemo "SQL" ="SELECT mid([Item No],1,2) AS FiberFamily, mid([Item No],6,1) AS FiberLetter, [Ba"
    "se]\015\012FROM [Basic Product Construction]\015\012WHERE [Basic Product Constru"
    "ction].[New Oracle Part #] not like \"*test*\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc9e6e260a74cfa4dbb6c91eddf18c4d4
End
dbMemo "Filter" ="((qryFiberOracle.Fiber2<>\"\"))"
dbMemo "OrderBy" ="FiberFamily"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="FiberFamily"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4b7f8aa1b0ac254192c15dcd6cf79c79
        End
    End
    Begin
        dbText "Name" ="FiberLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf311fa265539f24c80e5c0e9622e62e7
        End
    End
    Begin
        dbText "Name" ="[Base]"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x509b91749932c3418fc8d900c023de9e
        End
    End
End
