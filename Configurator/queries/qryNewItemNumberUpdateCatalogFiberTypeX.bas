Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"Fiber Type\") And ((tbl_Fibers.Fiber)="
    "Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType1))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="tblFiberElementValue"
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[SM/MM] & \" BIF \" & [ITU]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x1566563d2ccad949a101e449e5a564ea
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =376
    Right =1185
    Bottom =760
    Left =-1
    Top =-1
    Right =1077
    Bottom =170
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
        Left =306
        Top =18
        Right =535
        Bottom =121
        Top =0
        Name ="tblFiberElementValue"
        Name =""
    End
    Begin
        Left =573
        Top =6
        Right =817
        Bottom =169
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
