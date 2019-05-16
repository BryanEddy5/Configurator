dbMemo "SQL" ="SELECT qryFiberOracle.Oracle, tblBurnListingType.Type, tblBurnListingType.Listin"
    "gCompany, tblWOPExtract.RunTime, tblWOPExtract.Item, tblWOPExtract.ItemDesc, tbl"
    "WOPExtract.Subinv, tblWOPExtract.LotNumber, tblWOPExtract.Locator, tblWOPExtract"
    ".Qty, tblWOPExtract.ResQty, tblWOPExtract.UOM, tblWOPExtract.BOMFiberItem, tblWO"
    "PExtract.Wavelength, tblWOPExtract.LOTCreationDate\015\012FROM (tblWOPExtract IN"
    "NER JOIN qryFiberOracle ON tblWOPExtract.Item = qryFiberOracle.Oracle) INNER JOI"
    "N (tblBurnListing INNER JOIN tblBurnListingType ON (tblBurnListingType.Construct"
    "ion = tblBurnListing.ListingConstruction) AND (tblBurnListing.ListingSection = t"
    "blBurnListingType.Section) AND (tblBurnListingType.Tab = tblBurnListing.ListingT"
    "ab) AND (tblBurnListing.ListingCompany = tblBurnListingType.ListingCompany)) ON "
    "qryFiberOracle.Base = tblBurnListing.Base;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x2b9d47ed07d6394688a33a0fbad1873e
End
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.ItemDesc"
        dbInteger "ColumnWidth" ="12375"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Qty"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Subinv"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.ResQty"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.RunTime"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.LotNumber"
        dbInteger "ColumnWidth" ="2115"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.UOM"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.BOMFiberItem"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Wavelength"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Item"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Locator"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingType.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWOPExtract.LOTCreationDate"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingType.Type"
        dbLong "AggregateType" ="-1"
    End
End
