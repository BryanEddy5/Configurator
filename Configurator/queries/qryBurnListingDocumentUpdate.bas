Operation =1
Option =0
Where ="(((tblBurnListingDocument.ListingCompany)=[Forms]![frmNewListingUpdate]![Listing"
    "Company]))"
Begin InputTables
    Name ="tblBurnListingDocument"
End
Begin OutputColumns
    Expression ="tblBurnListingDocument.ListingDocument"
    Expression ="tblBurnListingDocument.ListingCompany"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcc3e4645396eb94c9d8e5b27b909de57
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBurnListingDocument.ListingDocument"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBurnListingDocument.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =319
    Top =340
    Right =1395
    Bottom =989
    Left =-1
    Top =-1
    Right =1044
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =346
        Bottom =261
        Top =0
        Name ="tblBurnListingDocument"
        Name =""
    End
End
