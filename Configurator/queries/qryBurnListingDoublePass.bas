dbMemo "SQL" ="SELECT [Basic (Cable/Sheath)].Active, tblBurnListing.ListingTab, tblBurnListing."
    "Base, tblBurnListing.ListingCompany, tblBurnListing.ListingSection, tblBurnListi"
    "ng.ListingConstruction, CDbl([Nominal OD])*39.37 AS OD, [Nom Wall]*39.37 AS NomW"
    "all, [Max Ave Wall]*39.37 AS MaxAveWall, [Min Ave Wall]*39.37 AS MinAveWall, [Mi"
    "n Spot Wall]*39.37 AS MinSpotWall, [Basic (Cable/Sheath)].[Jacket Material], Cou"
    "nt([Basic (Cable/Sheath)].Base) AS CountOfBase\015\012FROM (tblBurnListing INNER"
    " JOIN [Basic (Cable/Sheath)] ON tblBurnListing.Base = [Basic (Cable/Sheath)].Bas"
    "e) INNER JOIN qryItemNoSearch ON [Basic (Cable/Sheath)].[New Oracle Part Number]"
    " = qryItemNoSearch.Oracle\015\012GROUP BY [Basic (Cable/Sheath)].Active, tblBurn"
    "Listing.ListingTab, tblBurnListing.Base, tblBurnListing.ListingCompany, tblBurnL"
    "isting.ListingSection, tblBurnListing.ListingConstruction, CDbl([Nominal OD])*39"
    ".37, [Nom Wall]*39.37, [Max Ave Wall]*39.37, [Min Ave Wall]*39.37, [Min Spot Wal"
    "l]*39.37, [Basic (Cable/Sheath)].[Jacket Material]\015\012HAVING ((([Basic (Cabl"
    "e/Sheath)].Active)=-1) AND ((tblBurnListing.ListingConstruction) Is Null))\015\012"
    "ORDER BY tblBurnListing.ListingTab, tblBurnListing.Base, Count([Basic (Cable/She"
    "ath)].Base) DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf7e6f23d8316524ca9f11fc6ba86bf57
End
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblBurnListing.Base"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingCompany"
        dbInteger "ColumnWidth" ="1410"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingTab"
        dbInteger "ColumnWidth" ="1245"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingSection"
        dbInteger "ColumnWidth" ="1410"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingConstruction"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Active"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Jacket Material]"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8d98cce343f0ec44b5cb83aef64f122c
        End
    End
    Begin
        dbText "Name" ="NomWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfbf2027f5eab3a488690b8e44ad395a1
        End
    End
    Begin
        dbText "Name" ="MaxAveWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcd4f49394f5f1547acdeac3954a3d38e
        End
    End
    Begin
        dbText "Name" ="MinAveWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b1030b640aa1040b32300bfd03a0b5d
        End
    End
    Begin
        dbText "Name" ="MinSpotWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52ccdf2f4111d74798b79ea8d3522aaa
        End
    End
    Begin
        dbText "Name" ="CountOfBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8471f707840d54398dd36fce8ffbf36
        End
    End
End
