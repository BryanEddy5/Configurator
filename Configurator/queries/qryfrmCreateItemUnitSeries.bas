dbMemo "SQL" ="SELECT [Basic Product Construction].*, [Basic Product Construction].Active\015\012"
    "FROM ([Basic (Cable/Sheath)] INNER JOIN [Unit Series] ON [Basic (Cable/Sheath)]."
    "[Unit Series] = [Unit Series].[Unit Color Series]) INNER JOIN [Basic Product Con"
    "struction] ON [Unit Series].[Unit Type] = [Basic Product Construction].[Item No]"
    "\015\012WHERE ((([Basic Product Construction].Active)=-1) AND (([Basic (Cable/Sh"
    "eath)].[New Oracle Part Number])=\"PS07013-TEST\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6543c7b24646a144a7371a06bdd02481
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
