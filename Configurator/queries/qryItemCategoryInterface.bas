Operation =1
Option =0
Begin InputTables
    Name ="tblItemCategoryInterface"
End
Begin OutputColumns
    Expression ="tblItemCategoryInterface.*"
End
Begin OrderBy
    Expression ="tblItemCategoryInterface.CategorySetName"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x961a3cba62051d409e62b9316bf34c48
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategoryName"
        dbInteger "ColumnWidth" ="4815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =39
        Top =1
        Right =301
        Bottom =104
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
End
