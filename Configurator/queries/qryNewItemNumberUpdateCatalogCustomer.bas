Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"CUSTOMER VERSION\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![Customer]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd5b047c34a1d0d479529821bc8c621ed
End
Begin
End
Begin
    State =0
    Left =638
    Top =13
    Right =1747
    Bottom =337
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =124
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
End
