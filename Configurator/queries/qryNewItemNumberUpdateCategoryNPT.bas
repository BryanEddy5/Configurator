Operation =4
Option =8
Where ="(((tblItemCategoryInterface.CategorySetName)=\"New Product Type\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
    Name ="tblCableConstructionReferences"
    Name ="tblCableFamily"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="Nz([NewProductCode],[CategoryName])"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableFamily"
    Expression ="tblCableConstructionReferences.FamilyID = tblCableFamily.FamilyID"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
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
    0x24de1f9217b67e459d37f9d1900701ef
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
        dbText "Name" ="tblCableFamily.NewProductCode"
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
    Bottom =219
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =1055
        Top =42
        Right =1278
        Bottom =205
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
    Begin
        Left =494
        Top =44
        Right =638
        Bottom =188
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =129
        Top =56
        Right =420
        Bottom =278
        Top =0
        Name ="tblCableFamily"
        Name =""
    End
    Begin
        Left =759
        Top =39
        Right =903
        Bottom =240
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
