dbMemo "SQL" ="SELECT tblFiberSpecs.*\015\012FROM qryFiberSpecsMax INNER JOIN tblFiberSpecs ON "
    "(qryFiberSpecsMax.MaxOfFiberTypeCount=tblFiberSpecs.FiberTypeCount) AND (qryFibe"
    "rSpecsMax.FiberType=tblFiberSpecs.FiberType)\015\012WHERE (((qryFiberSpecsMax.Fi"
    "berType) Like \"fbr0*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0943556197b782489b592fb7313bff1a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblFiberSpecs.FiberTypeCount"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFiberSpecs.SpecificationElement"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
