Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"DIAMETER\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="getCatalogOD([Item No])"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblItemCatalogInterface"
    Expression ="[Basic Product Construction].[New Oracle Part #] = tblItemCatalogInterface.ItemN"
        "umber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x5d510acc440f0b4bbcb2175c8ae7d205
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
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
    Bottom =225
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =267
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
    Begin
        Left =341
        Top =2
        Right =642
        Bottom =217
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
