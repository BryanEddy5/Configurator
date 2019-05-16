dbMemo "SQL" ="SELECT BasicProductConstruction_LIVE.Active, tblBurnListing.ListingTab, tblBurnL"
    "isting.Base, tblBurnListing.ListingCompany, tblBurnListing.ListingSection, tblBu"
    "rnListing.ListingConstruction, CDbl([Nominal OD])*39.37 AS OD, [Nominal Wall]*39"
    ".37 AS NomWall, [Max Ave Wall]*39.37 AS MaxAveWall, [Min Ave Wall]*39.37 AS MinA"
    "veWall, [Min Spot Wall]*39.37 AS MinSpotWall, BasicProductConstruction_LIVE.[Jac"
    "ket Material], Count(BasicProductConstruction_LIVE.Base) AS CountOfBase, BasicPr"
    "oductConstruction_LIVE.[TB Material]\015\012FROM qryItemNoSearch INNER JOIN (Bas"
    "icProductConstruction_LIVE INNER JOIN tblBurnListing ON BasicProductConstruction"
    "_LIVE.Base = tblBurnListing.Base) ON qryItemNoSearch.Oracle = BasicProductConstr"
    "uction_LIVE.[New Oracle Part #]\015\012GROUP BY BasicProductConstruction_LIVE.Ac"
    "tive, tblBurnListing.ListingTab, tblBurnListing.Base, tblBurnListing.ListingComp"
    "any, tblBurnListing.ListingSection, tblBurnListing.ListingConstruction, CDbl([No"
    "minal OD])*39.37, [Nominal Wall]*39.37, [Max Ave Wall]*39.37, [Min Ave Wall]*39."
    "37, [Min Spot Wall]*39.37, BasicProductConstruction_LIVE.[Jacket Material], Basi"
    "cProductConstruction_LIVE.[TB Material]\015\012HAVING (((BasicProductConstructio"
    "n_LIVE.Active)=-1))\015\012ORDER BY tblBurnListing.ListingTab, tblBurnListing.Ba"
    "se, Count(BasicProductConstruction_LIVE.Base) DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb39a392730d1174894a1284fae145f59
End
dbMemo "OrderBy" ="qryBurnListingSinglePass.Base"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="BasicProductConstruction_LIVE.Active"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.Base"
        dbInteger "ColumnOrder" ="2"
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
        dbInteger "ColumnOrder" ="3"
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
        dbText "Name" ="BasicProductConstruction_LIVE.[Jacket Material]"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5ed7f05dd65e5941872384e7d8ba04cb
        End
    End
    Begin
        dbText "Name" ="NomWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd78e0d81ba7a6428971cd2da33e88a8
        End
    End
    Begin
        dbText "Name" ="MaxAveWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19a7347315a0014385334a327d880c1e
        End
    End
    Begin
        dbText "Name" ="MinAveWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xae3da7b2abecc949bba06d3dfa12cf3f
        End
    End
    Begin
        dbText "Name" ="MinSpotWall"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa3f1a8fb90d2c34a8f0a127912e015d2
        End
    End
    Begin
        dbText "Name" ="CountOfBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x821a365f7a84674f82666a21c2b9af9c
        End
    End
End
