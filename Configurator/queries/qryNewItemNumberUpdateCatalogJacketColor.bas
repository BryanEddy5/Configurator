Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"JACKET COLOR\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetails.Form!JacketColor"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x276368720916664e8543e7d660e54a6e
End
Begin
End
Begin
    State =0
    Left =61
    Top =17
    Right =1170
    Bottom =341
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
