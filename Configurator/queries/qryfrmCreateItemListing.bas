Operation =1
Option =2
Where ="(((tblBurnListing.Base)=[Forms]![frmCreateItem]![Base]))"
Begin InputTables
    Name ="tblBurnListing"
End
Begin OutputColumns
    Expression ="tblBurnListing.ListingCompany"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcd9160040f8f0f4da809e7f4ec39371c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =190
    Top =182
    Right =1565
    Bottom =617
    Left =-1
    Top =-1
    Right =1343
    Bottom =204
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =169
        Top =0
        Name ="tblBurnListing"
        Name =""
    End
End
