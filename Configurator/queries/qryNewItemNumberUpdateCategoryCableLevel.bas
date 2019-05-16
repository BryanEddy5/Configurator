Operation =4
Option =0
Where ="(((tblItemCategoryInterface.CategorySetName)=\"Cable Level\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
    Name ="Basic Product Construction"
    Name ="tblCableConstructionReferences"
    Name ="tblDesignTypes"
    Name ="Oracle_CableLevel"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="[Oracle_CableLevel].[CableLevel]"
End
Begin Joins
    LeftTable ="tblItemCategoryInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblItemCategoryInterface.ItemNumber = [Basic Product Construction].[New Oracle P"
        "art #]"
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
    RightTable ="Oracle_CableLevel"
    Expression ="tblDesignTypes.CableLevelGuid = Oracle_CableLevel.CableLevelGuid"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xe545adc5b0299c428376d17f2def5d92
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
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1197
    Bottom =775
    Left =-1
    Top =-1
    Right =1077
    Bottom =219
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
    Begin
        Left =350
        Top =34
        Right =769
        Bottom =178
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =182
        Top =47
        Right =592
        Bottom =263
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
    Begin
        Left =18
        Top =23
        Right =162
        Bottom =167
        Top =0
        Name ="Oracle_CableLevel"
        Name =""
    End
End
