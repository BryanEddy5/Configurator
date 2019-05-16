dbMemo "SQL" ="SELECT tblCorningPrint.CorningPrint, tblFiberExtensions.Fiber\015\012FROM tblFib"
    "erExtensions INNER JOIN tblCorningPrint ON tblFiberExtensions.FiberExtension = t"
    "blCorningPrint.FiberExtension\015\012WHERE (((tblFiberExtensions.Fiber)=[Forms]!"
    "[frmCreateItem]![Fiber1]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x79c86130528aef49ab33d56bda4e5911
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCorningPrint.CorningPrint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFiberExtensions.Fiber"
        dbLong "AggregateType" ="-1"
    End
End
