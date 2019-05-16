dbMemo "SQL" ="INSERT INTO tblUnitSeriesCopy\015\012SELECT [Unit Series].*\015\012FROM (qryCopy"
    "ItemMatch INNER JOIN [Basic (Cable/Sheath)] ON qryCopyItemMatch.Oracle = [Basic "
    "(Cable/Sheath)].[New Oracle Part Number]) INNER JOIN [Unit Series] ON [Basic (Ca"
    "ble/Sheath)].[Unit Series] = [Unit Series].[Unit Color Series];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6e564bc8462d1845893cdcdf6ca3b56d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
