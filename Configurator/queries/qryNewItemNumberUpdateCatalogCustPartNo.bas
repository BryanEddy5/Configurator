Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"CUSTOMER PART NUMBER\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[Forms]![frmItemInterface]![frmPremiseItemDetails].[form].[CustomerPartNumber] &"
        " \" REV \" & [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![Customer"
        "Rev]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x450a0e31cf1aa24285cdea2185072091
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =128
    Right =1215
    Bottom =757
    Left =-1
    Top =-1
    Right =1077
    Bottom =59
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
