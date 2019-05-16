Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName) Like \"*cable type*\") And ((tblCatalogC"
    "ableType.Prefix)=Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleIt"
    "emNumber,2)))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="tblCatalogCableType"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[CableType]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa40b3ad465191a41b9a2325c563f389e
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="4125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =57
    Top =223
    Right =1166
    Bottom =746
    Left =-1
    Top =-1
    Right =1077
    Bottom =309
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =251
        Top =55
        Right =543
        Bottom =203
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
    Begin
        Left =733
        Top =113
        Right =829
        Bottom =201
        Top =0
        Name ="tblCatalogCableType"
        Name =""
    End
End
