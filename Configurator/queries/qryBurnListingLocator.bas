Operation =1
Option =2
Begin InputTables
    Name ="tblBurnListingConstruction"
    Name ="tblWOPExtract"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="dbo_CableUnion"
    Name ="tblBurnListing"
    Name ="AFLPRD_INVSysItemCost_CAB"
    Name ="qryBasePrint"
    Name ="tblBurnListingRating"
End
Begin OutputColumns
    Expression ="dbo_CableUnion.Oracle"
    Expression ="dbo_CableUnion.[Item No]"
    Expression ="tblWOPExtract.Locator"
    Alias ="Expr1"
    Expression ="tblWOPExtract.Qty"
    Expression ="tblWOPExtract.LotNumber"
    Expression ="tblWOPExtract.Subinv"
    Expression ="tblWOPExtract.UOM"
    Expression ="tblWOPExtract.ItemDesc"
    Alias ="NomODInches"
    Expression ="[NominalOD]*0.03936996"
    Alias ="BurnLengthFT_OFNP_ETL"
    Expression ="Round((12/[NomODInches])*24*1.1)"
    Alias ="BurnLengthFT_OFNR_ETL"
    Expression ="Round((12/[NomODInches])*16.5*1.1)"
    Alias ="BurnLengthFT_OFNP_UL"
    Expression ="Round((900/([NomODInches]))*1.1/3)"
    Alias ="QtyFt"
    Expression ="Round([Qty]*3.28)"
    Expression ="AFLPRD_INVSysItemCost_CAB.Item_Status"
    Expression ="qryBasePrint.PrintLine3"
    Expression ="AFLPRD_INVSysItemCost_CAB.Item_Cost"
    Expression ="tblBurnListing.ListingTab"
    Expression ="tblBurnListing.ListingSection"
    Expression ="tblBurnListing.ListingConstruction"
    Expression ="tblBurnListingRating.Rating"
    Expression ="tblBurnListing.ListingCompany"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="dbo_CableUnion"
    Expression ="tblCableConstructionReferences.Base = dbo_CableUnion.Base"
    Flag =1
    LeftTable ="dbo_CableUnion"
    RightTable ="tblBurnListing"
    Expression ="dbo_CableUnion.ListingCompany = tblBurnListing.ListingCompany"
    Flag =1
    LeftTable ="dbo_CableUnion"
    RightTable ="tblBurnListing"
    Expression ="dbo_CableUnion.Base = tblBurnListing.Base"
    Flag =1
    LeftTable ="tblWOPExtract"
    RightTable ="dbo_CableUnion"
    Expression ="tblWOPExtract.Item = dbo_CableUnion.Oracle"
    Flag =1
    LeftTable ="dbo_CableUnion"
    RightTable ="AFLPRD_INVSysItemCost_CAB"
    Expression ="dbo_CableUnion.Oracle = AFLPRD_INVSysItemCost_CAB.ItemNumber"
    Flag =1
    LeftTable ="dbo_CableUnion"
    RightTable ="qryBasePrint"
    Expression ="dbo_CableUnion.Oracle = qryBasePrint.Oracle"
    Flag =1
    LeftTable ="tblBurnListingConstruction"
    RightTable ="tblBurnListing"
    Expression ="tblBurnListingConstruction.ListingSection = tblBurnListing.ListingSection"
    Flag =1
    LeftTable ="tblBurnListingConstruction"
    RightTable ="tblBurnListing"
    Expression ="tblBurnListingConstruction.ListingDocument = tblBurnListing.ListingTab"
    Flag =1
    LeftTable ="tblBurnListingConstruction"
    RightTable ="tblBurnListing"
    Expression ="tblBurnListingConstruction.ListingCompany = tblBurnListing.ListingCompany"
    Flag =1
    LeftTable ="tblBurnListingConstruction"
    RightTable ="tblBurnListingRating"
    Expression ="tblBurnListingConstruction.ConstructionID = tblBurnListingRating.ConstructionID"
    Flag =1
End
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
    0x9fa9ead35ef68948b08a3964728f7bd7
End
Begin
    Begin
        dbText "Name" ="tblWOPExtract.Locator"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="tblWOPExtract.UOM"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="QtyFt"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e8eeea2efbb774ab5af5fb2b3ef4b4a
        End
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="tblWOPExtract.ItemDesc"
        dbInteger "ColumnWidth" ="4695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="tblWOPExtract.LotNumber"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="tblWOPExtract.Subinv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Oracle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.[Item No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="NomODInches"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d884c21e5471e4695d7d9d1f5935ea0
        End
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="BurnLengthFT_OFNP_UL"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="BurnLengthFT_OFNP_ETL"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="qryBasePrint.PrintLine3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="BurnLengthFT_OFNR_ETL"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingTab"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingSection"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingCompany"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingConstruction"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb502559c99dafe4da9fa1c6b98bf95f3
        End
    End
End
Begin
    State =0
    Left =59
    Top =84
    Right =1607
    Bottom =743
    Left =-1
    Top =-1
    Right =1516
    Bottom =137
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblBurnListingConstruction"
        Name =""
    End
    Begin
        Left =268
        Top =23
        Right =621
        Bottom =461
        Top =0
        Name ="tblWOPExtract"
        Name =""
    End
    Begin
        Left =1057
        Top =31
        Right =1201
        Bottom =175
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1274
        Top =16
        Right =1505
        Bottom =425
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =58
        Top =105
        Right =261
        Bottom =327
        Top =0
        Name ="dbo_CableUnion"
        Name =""
    End
    Begin
        Left =641
        Top =12
        Right =785
        Bottom =300
        Top =0
        Name ="tblBurnListing"
        Name =""
    End
    Begin
        Left =833
        Top =12
        Right =977
        Bottom =156
        Top =0
        Name ="AFLPRD_INVSysItemCost_CAB"
        Name =""
    End
    Begin
        Left =717
        Top =293
        Right =861
        Bottom =437
        Top =0
        Name ="qryBasePrint"
        Name =""
    End
    Begin
        Left =909
        Top =156
        Right =1053
        Bottom =300
        Top =0
        Name ="tblBurnListingRating"
        Name =""
    End
End
