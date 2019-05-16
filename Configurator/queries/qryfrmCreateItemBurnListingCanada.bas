Operation =1
Option =0
Where ="(((vBurnListingBase.ListingCompany)=fORMS!frmCreateItem!ListingCompany) And ((vB"
    "urnListingBase.Base)=Forms!frmCreateItem!Base) And ((vBurnListingBase.Canada)<>0"
    ") And ((tblBurnListingRatingCombinations.RatingUS)=Forms!frmCreateItem!ListingTy"
    "pe))"
Begin InputTables
    Name ="vBurnListingBase"
    Name ="tblBurnListingRatingCombinations"
End
Begin OutputColumns
    Expression ="vBurnListingBase.Rating"
    Expression ="vBurnListingBase.ListingConstruction"
    Expression ="vBurnListingBase.ListingSection"
    Expression ="vBurnListingBase.ListingTab"
    Expression ="vBurnListingBase.ListingCompany"
    Expression ="vBurnListingBase.Base"
    Expression ="vBurnListingBase.US"
    Expression ="vBurnListingBase.Canada"
    Expression ="tblBurnListingRatingCombinations.RatingUS"
End
Begin Joins
    LeftTable ="vBurnListingBase"
    RightTable ="tblBurnListingRatingCombinations"
    Expression ="vBurnListingBase.Rating = tblBurnListingRatingCombinations.RatingCanada"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x379f63e8f1e0954f8471759b1a9454b2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vBurnListingBase.ListingConstruction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingTab"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.US"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingSection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Canada"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Rating"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingRatingCombinations.RatingUS"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =25
    Top =97
    Right =1573
    Bottom =985
    Left =-1
    Top =-1
    Right =1516
    Bottom =130
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =346
        Top =33
        Right =660
        Bottom =194
        Top =0
        Name ="vBurnListingBase"
        Name =""
    End
    Begin
        Left =763
        Top =29
        Right =907
        Bottom =173
        Top =0
        Name ="tblBurnListingRatingCombinations"
        Name =""
    End
End
