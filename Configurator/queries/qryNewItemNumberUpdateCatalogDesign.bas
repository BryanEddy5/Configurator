Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"DESIGN CODE\" Or (tblItemCatalogInterf"
    "ace.ElementName)=\"Subassembly Type\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0x00b267bab4d8b5449d1f6e562f3cf45f
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
    Left =109
    Top =70
    Right =1218
    Bottom =394
    Left =-1
    Top =-1
    Right =1077
    Bottom =76
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
