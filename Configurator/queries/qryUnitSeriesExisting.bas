dbMemo "SQL" ="SELECT DISTINCT [Basic (Cable/Sheath)].[Unit Series]\015\012FROM [Basic (Cable/S"
    "heath)] INNER JOIN [Unit Series] ON [Basic (Cable/Sheath)].[Unit Series]=[Unit S"
    "eries].[Unit Color Series]\015\012GROUP BY [Basic (Cable/Sheath)].[Unit Series],"
    " [Basic (Cable/Sheath)].[Item No]\015\012HAVING ((([Basic (Cable/Sheath)].[Unit "
    "Series]) Is Not Null) And (([Basic (Cable/Sheath)].[Item No]) Like Left(Forms!fr"
    "mMainMenu!ItemNo,5) & \"?\" & Mid(Forms!frmMainMenu!ItemNo,7,2) & \"????\" & \"*"
    "\"))\015\012ORDER BY [Basic (Cable/Sheath)].[Unit Series];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb4a2bbe8aaba3840a6770651a002353f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Unit Series]"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
