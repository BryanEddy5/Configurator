dbMemo "SQL" ="SELECT [Basic (Cable/Sheath)].[New Oracle Part Number], [Basic (Cable/Sheath)].["
    "Item No], [Basic (Cable/Sheath)].[Jacket Material], [Basic (Cable/Sheath)].[Jack"
    "et Color], \"N/A\" AS ColorChip, Left([Fiber Type A],8) AS Fiber\015\012FROM qry"
    "MainMenuMatch INNER JOIN [Basic (Cable/Sheath)] ON qryMainMenuMatch.Oracle = [Ba"
    "sic (Cable/Sheath)].[New Oracle Part Number];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfe9f669fffb26d48b2ebf320cd346ec8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb689ea279adc8c4490840bb402cc4bb8
        End
    End
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c5618e64b98fc4fae629076cbcf6da6
        End
    End
End
