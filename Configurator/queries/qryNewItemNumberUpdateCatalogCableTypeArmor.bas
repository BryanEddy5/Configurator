Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName) Like \"*cable type*\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[ElementValue] & \" - ARMORED\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x4bc276db307fce489efd8305ee037c5c
End
Begin
End
Begin
    State =0
    Left =-8
    Top =68
    Right =1101
    Bottom =591
    Left =-1
    Top =-1
    Right =1077
    Bottom =309
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =428
        Top =64
        Right =720
        Bottom =212
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
End
