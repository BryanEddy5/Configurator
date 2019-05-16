Operation =4
Option =0
Where ="(((tblItemCategoryInterface.CategorySetName)=\"Product Class\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
    Name ="tbl_Fibers"
    Name ="Basic Product Construction"
    Name ="tblCableConstructionReferences"
    Name ="tblDesignTypeGroups"
    Name ="tblDesignTypes"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="UpdateProductClass([Oracle_Product_Class_3],[Oracle_Product_Class_5],[Mode],[Fib"
        "erType2],[New Oracle Part #],[Item No])"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblItemCategoryInterface"
    Expression ="[Basic Product Construction].[New Oracle Part #] = tblItemCategoryInterface.Item"
        "Number"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tbl_Fibers"
    Expression ="[Basic Product Construction].[Fiber Oracle item] = tbl_Fibers.Fiber"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignTypes"
    Expression ="tblCableConstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID"
    Flag =1
    LeftTable ="tblDesignTypes"
    RightTable ="tblDesignTypeGroups"
    Expression ="tblDesignTypes.DesignTypeGroupID = tblDesignTypeGroups.DesignTypeGroupID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb32f744a3dadb042b8dbee571cf38955
End
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategoryName"
        dbInteger "ColumnWidth" ="4860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3450"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =-448
    Top =164
    Right =661
    Bottom =793
    Left =-1
    Top =-1
    Right =1077
    Bottom =269
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =610
        Top =17
        Right =848
        Bottom =188
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =236
        Bottom =244
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
    Begin
        Left =302
        Top =8
        Right =588
        Bottom =249
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =872
        Top =55
        Right =1111
        Bottom =257
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1298
        Top =15
        Right =1476
        Bottom =208
        Top =0
        Name ="tblDesignTypeGroups"
        Name =""
    End
    Begin
        Left =1069
        Top =14
        Right =1277
        Bottom =211
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
End
