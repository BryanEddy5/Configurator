Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"CUSTOMER VERSION\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form!Customer"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xea2a275a2e12fc48b59dd039a51ae4e9
End
Begin
End
Begin
    State =0
    Left =104
    Top =98
    Right =1396
    Bottom =422
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
