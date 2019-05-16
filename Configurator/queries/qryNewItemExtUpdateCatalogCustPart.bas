Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"CUSTOMER PART NUMBER\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form!CustomerPartNumber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x6ced3c1d33bde7439c69757a017465b1
End
Begin
End
Begin
    State =0
    Left =61
    Top =69
    Right =1353
    Bottom =393
    Left =-1
    Top =-1
    Right =1260
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
