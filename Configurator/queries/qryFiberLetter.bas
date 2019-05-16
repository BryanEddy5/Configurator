dbMemo "SQL" ="SELECT Fiber, Fiber_Type as FiberLetter\015\012 FROM qryFibers\015\012\015\012\015"
    "\012UNION\015\012SELECT Fiber, Alt1_Fiber_Type as FiberLetter\015\012FROM qryFib"
    "ers\015\012\015\012UNION\015\012SELECT Fiber, Alt2_Fiber_Type as FiberLetter\015"
    "\012FROM qryFibers\015\012\015\012UNION\015\012SELECT Fiber,Alt3_Fiber_Type as F"
    "iberLetter\015\012FROM qryFibers\015\012\015\012UNION SELECT Fiber, Alt4_Fiber_T"
    "ype as FiberLetter\015\012FROM qryFibers;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc1546a6d2a563f458b438228060ed4a3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xec0076ac58512845a6f5e3a65b582a6c
        End
    End
    Begin
        dbText "Name" ="FiberLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8fe95e65a2d14348ac203d3d5fa2d9c2
        End
    End
End
