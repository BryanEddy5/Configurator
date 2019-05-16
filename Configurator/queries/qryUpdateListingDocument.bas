Operation =4
Option =0
Where ="(((tblBurnListingDocument.ListingDocument)=Forms!frmNewListingUpdate!ReplaceDocu"
    "ment) And ((tblBurnListingDocument.ListingCompany)=Forms!frmNewListingUpdate!Lis"
    "tingCompany))"
Begin InputTables
    Name ="tblBurnListingDocument"
End
Begin OutputColumns
    Name ="tblBurnListingDocument.ListingDocument"
    Expression ="[Forms]![frmNewListingUpdate]![Listing]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x0c10c99e6524b44e983ac353c515c33a
End
Begin
    Begin
        dbText "Name" ="tblBurnListingDocument.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingDocument.ListingDocument"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =408
    Top =78
    Right =1517
    Bottom =707
    Left =-1
    Top =-1
    Right =1077
    Bottom =398
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =298
        Top =45
        Right =677
        Bottom =254
        Top =0
        Name ="tblBurnListingDocument"
        Name =""
    End
End
