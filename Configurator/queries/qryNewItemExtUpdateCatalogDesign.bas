Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"DESIGN CODE\" Or (tblItemCatalogInterf"
    "ace.ElementName)=\"subassembly type\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form!DesignCode"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x28f9efe8f363024ca10782e4185ef6a8
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-517
    Top =479
    Right =775
    Bottom =803
    Left =-1
    Top =-1
    Right =1260
    Bottom =93
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
