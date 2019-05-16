dbMemo "SQL" ="SELECT DISTINCT [Basic Product Construction].Base, [Listing Company]\015\012FROM"
    " [Basic Product Construction]\015\012WHERE [Basic Product Construction].[Active]"
    " <> 0 AND  (([Basic Product Construction].[New Oracle Part #])  not like \"*test"
    "*\") and  [Print Line 3] <> \"\";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb27d3bd9a81bad48964b04e1a7333244
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Listing Company]"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Listing Company"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55bb742f0a805546aac959eeb7747a39
        End
    End
End
