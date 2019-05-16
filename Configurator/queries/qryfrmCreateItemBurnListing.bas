Operation =1
Option =0
Begin InputTables
    Name ="tblBurnListing"
    Name ="tblBurnListingConstruction"
    Name ="tblBurnListingRating"
    Name ="vBurnListingBase"
End
Begin OutputColumns
    Expression ="tblBurnListing.ListingCompany"
    Expression ="tblBurnListing.Base"
    Expression ="tblBurnListing.ListingTab"
    Expression ="tblBurnListing.ListingSection"
    Expression ="tblBurnListing.ListingConstruction"
    Expression ="tblBurnListingConstruction.ConstructionID"
    Expression ="tblBurnListingRating.Rating"
    Expression ="tblBurnListingRating.US"
    Expression ="tblBurnListingRating.Canada"
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
    0xd2a8b790c5fb38408d47c74a9a52dbff
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBurnListing.ListingSection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingConstruction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingRating.US"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingRating.Canada"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListing.ListingTab"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingConstruction.ConstructionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingRating.Rating"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =87
    Top =78
    Right =1354
    Bottom =643
    Left =-1
    Top =-1
    Right =1235
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =235
        Top =35
        Right =443
        Bottom =177
        Top =0
        Name ="tblBurnListing"
        Name =""
    End
    Begin
        Left =598
        Top =10
        Right =820
        Bottom =163
        Top =0
        Name ="tblBurnListingConstruction"
        Name =""
    End
    Begin
        Left =898
        Top =23
        Right =1042
        Bottom =183
        Top =0
        Name ="tblBurnListingRating"
        Name =""
    End
    Begin
        Left =44
        Top =61
        Right =188
        Bottom =205
        Top =0
        Name ="vBurnListingBase"
        Name =""
    End
End
