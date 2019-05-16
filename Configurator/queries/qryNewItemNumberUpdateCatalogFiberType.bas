Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"Fiber Type\" Or (tblItemCatalogInterfa"
    "ce.ElementName)=\"Fiber Type 1\") And ((Mid(Forms!frmItemInterface!frmPremiseIte"
    "mDetails.Form!DesignCode,6,1))=[FiberLetter]))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="tblFiberElementValue"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[FiberValue]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb592a1be286c264a8322ae6c3be5e1b1
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-60
    Top =151
    Right =1049
    Bottom =780
    Left =-1
    Top =-1
    Right =1077
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =154
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
    Begin
        Left =403
        Top =12
        Right =561
        Bottom =115
        Top =0
        Name ="tblFiberElementValue"
        Name =""
    End
End
