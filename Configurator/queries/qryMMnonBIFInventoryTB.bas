dbMemo "SQL" ="SELECT qryFiberOracle.Oracle, qryFiberOracle.[Item No], qryFiberOracle.Fiber, Fi"
    "_fiber_length_extract_AFLPRD.[Onhand Qty], Fi_fiber_length_extract_AFLPRD.[Lot N"
    "umber], Fi_fiber_length_extract_AFLPRD.Subinventory, Fi_fiber_length_extract_AFL"
    "PRD.Locator\015\012FROM Fi_fiber_length_extract_AFLPRD INNER JOIN qryFiberOracle"
    " ON Fi_fiber_length_extract_AFLPRD.[Item Number]=qryFiberOracle.Oracle\015\012WH"
    "ERE (((qryFiberOracle.Oracle) Not Like \"fb*\") AND ((qryFiberOracle.Fiber) Like"
    " \"fbr00131*\" Or (qryFiberOracle.Fiber) Like \"fbr00119*\" Or (qryFiberOracle.F"
    "iber) Like \"fbr00161*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x873a2bac49652c46a856ae8a2d57d227
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnWidth" ="2475"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.BIF"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Fiber"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Mode"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.[OM#]"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.SpecialNotes"
        dbInteger "ColumnWidth" ="1695"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="3345"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.Item_Status"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Description_Oracle"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="4365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Subinv"
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Locator"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Qty"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.[Item Number]"
        dbInteger "ColumnWidth" ="1800"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.[Onhand Qty]"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.[Lot Number]"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.Subinventory"
        dbInteger "ColumnWidth" ="1785"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fi_fiber_length_extract_AFLPRD.Locator"
        dbInteger "ColumnWidth" ="2040"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
