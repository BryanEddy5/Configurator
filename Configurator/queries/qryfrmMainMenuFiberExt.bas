dbMemo "SQL" ="SELECT DISTINCT Mid([Forms]![frmMainMenu]![ItemNo],14,30) AS FiberExtension\015\012"
    "FROM tbl_Fibers INNER JOIN tblFiberExtensions ON tbl_Fibers.Fiber = tblFiberExte"
    "nsions.Fiber\015\012WHERE (((Mid([Forms]![frmMainMenu]![ItemNo],14,30)) Like \"*"
    "\" & RTrim([FiberExtension]) & \"*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe0d812b7ab5f8f49bcd5c6215914e2a2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FiberExtension"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
End
