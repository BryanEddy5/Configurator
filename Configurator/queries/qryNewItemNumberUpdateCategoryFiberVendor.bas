Operation =4
Option =0
Where ="(((tblItemCategoryInterface.CategorySetName)=\"Fiber Vendor\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="IIf([VendorSpecificFiber]=False,\"Generic\",\"Specific\")"
End
Begin Joins
    LeftTable ="tblItemCategoryInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblItemCategoryInterface.ItemNumber = [Basic Product Construction].[New Oracle P"
        "art #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb5abf5a09119e644bd9f7abda874883e
End
Begin
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategoryName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategorySetName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].VendorSpecificFiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =308
        Top =6
        Right =661
        Bottom =263
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
    Begin
        Left =839
        Top =17
        Right =1392
        Bottom =246
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
