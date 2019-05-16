Operation =5
Option =0
Where ="(((tblItemCategoryInterface.CategoryName) Is Null) AND ((tblItemCategoryInterfac"
    "e.CategorySetName)<>\"product Class\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
End
Begin OutputColumns
    Expression ="tblItemCategoryInterface.*"
    Expression ="tblItemCategoryInterface.CategoryName"
    Expression ="tblItemCategoryInterface.CategorySetName"
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
    0x91fa276b6c9ec34e9cd042a0631b7e8c
End
Begin
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategoryName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4440"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategorySetName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.AllowOverride"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.LimitToList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =353
    Right =1657
    Bottom =966
    Left =-1
    Top =-1
    Right =1506
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =38
        Top =6
        Right =226
        Bottom =207
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
End
