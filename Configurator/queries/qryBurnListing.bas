Operation =1
Option =0
Begin InputTables
    Name ="tblBurnListing"
    Name ="tblBurnListingConstruction"
    Name ="tblBurnListingRating"
End
Begin OutputColumns
    Expression ="tblBurnListing.ListingCompany"
    Expression ="tblBurnListing.ListingTab"
    Expression ="tblBurnListing.ListingSection"
    Expression ="tblBurnListing.ListingConstruction"
    Expression ="tblBurnListing.Base"
    Expression ="tblBurnListingRating.Rating"
End
Begin Joins
    LeftTable ="tblBurnListing"
    RightTable ="tblBurnListingConstruction"
    Expression ="tblBurnListing.ListingCompany = tblBurnListingConstruction.ListingCompany"
    Flag =1
    LeftTable ="tblBurnListing"
    RightTable ="tblBurnListingConstruction"
    Expression ="tblBurnListing.ListingTab = tblBurnListingConstruction.ListingDocument"
    Flag =1
    LeftTable ="tblBurnListing"
    RightTable ="tblBurnListingConstruction"
    Expression ="tblBurnListing.ListingSection = tblBurnListingConstruction.ListingSection"
    Flag =1
    LeftTable ="tblBurnListing"
    RightTable ="tblBurnListingConstruction"
    Expression ="tblBurnListing.ListingConstruction = tblBurnListingConstruction.ListingConstruct"
        "ion"
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
dbBinary "GUID" = Begin
    0x6a61219c30aecb4b88fced020923f4e9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBurnListing.ListingCompany"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingTab"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingSection"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingConstruction"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingRating.Rating"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =11
    Top =420
    Right =1789
    Bottom =1274
    Left =-1
    Top =-1
    Right =1746
    Bottom =152
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =280
        Top =14
        Right =473
        Bottom =201
        Top =0
        Name ="tblBurnListing"
        Name =""
    End
    Begin
        Left =535
        Top =12
        Right =679
        Bottom =156
        Top =0
        Name ="tblBurnListingConstruction"
        Name =""
    End
    Begin
        Left =785
        Top =8
        Right =929
        Bottom =152
        Top =0
        Name ="tblBurnListingRating"
        Name =""
    End
End
