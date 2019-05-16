Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName) Like \"*cable rating*\") And ((tblCatalo"
    "gCableRating.JacketLetter)=Mid(Forms!frmItemInterface!frmPremiseItemDetails.Form"
    "!DesignCode,2,1)))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="tblCatalogCableRating"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[CableRating]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x899f7887b646834f8257c738ca4865ad
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
        Left =852
        Top =26
        Right =1148
        Bottom =114
        Top =0
        Name ="tblCatalogCableRating"
        Name =""
    End
End
