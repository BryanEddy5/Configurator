Operation =1
Option =0
Where ="(((qryFiberOracle.Oracle) Not Like \"fb*\") AND ((qryFiberOracle.Fiber) Like \"f"
    "br00131*\" Or (qryFiberOracle.Fiber) Like \"fbr00119*\" Or (qryFiberOracle.Fiber"
    ") Like \"fbr00161*\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tblWOPExtract"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Fiber"
    Expression ="tblWOPExtract.Subinv"
    Expression ="tblWOPExtract.LotNumber"
    Expression ="tblWOPExtract.Locator"
    Alias ="Expr1"
    Expression ="tblWOPExtract.Qty"
    Expression ="qryFiberOracle.PID"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="tblWOPExtract"
    Expression ="qryFiberOracle.Oracle = tblWOPExtract.Item"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf5269a2e873d0744b0e9a551a2ca9bdf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
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
        dbInteger "ColumnWidth" ="1245"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Locator"
        dbInteger "ColumnWidth" ="2130"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Qty"
        dbInteger "ColumnWidth" ="2715"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="7"
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
    Begin
        dbText "Name" ="tblWOPExtract.LotNumber"
        dbInteger "ColumnWidth" ="2940"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3d7a1fe439cb8849930bf89a1ea1300b
        End
    End
End
Begin
    State =0
    Left =35
    Top =190
    Right =1327
    Bottom =671
    Left =-1
    Top =-1
    Right =1260
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =53
        Top =13
        Right =394
        Bottom =206
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =496
        Top =52
        Right =925
        Bottom =279
        Top =0
        Name ="tblWOPExtract"
        Name =""
    End
End
