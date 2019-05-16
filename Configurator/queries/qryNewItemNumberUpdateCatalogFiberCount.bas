Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"FIBER COUNT\" Or (tblItemCatalogInterf"
    "ace.ElementName)=\"Fiber Count 1\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="Int(Mid(Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode,3,3))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xaae97f89c53b564bbe0059b0dd28c473
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
    Left =-731
    Top =92
    Right =378
    Bottom =416
    Left =-1
    Top =-1
    Right =1077
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
