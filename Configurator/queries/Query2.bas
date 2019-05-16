Operation =1
Option =0
Begin InputTables
    Name ="tblGetWhereUsedMaterial"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="tblGetWhereUsedMaterial.Material"
    Expression ="tblGetWhereUsedMaterial.AssemblyItemNumber"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="tblGetWhereUsedMaterial.MaterialQuantity_PerPrimaryUOM"
    Expression ="tblGetWhereUsedMaterial.MaterialQuanitty_PerOrder"
    Expression ="tblGetWhereUsedMaterial.PrimaryUOM"
    Expression ="tblGetWhereUsedMaterial.Item_Status"
    Expression ="tblGetWhereUsedMaterial.Customer"
    Expression ="tblGetWhereUsedMaterial.OrderNum"
    Expression ="tblGetWhereUsedMaterial.Shipped"
    Expression ="tblGetWhereUsedMaterial.Quantity"
    Expression ="tblGetWhereUsedMaterial.QTY_UOM"
    Expression ="tblGetWhereUsedMaterial.CountPerUOM"
End
Begin Joins
    LeftTable ="tblGetWhereUsedMaterial"
    RightTable ="Basic Product Construction"
    Expression ="tblGetWhereUsedMaterial.AssemblyItemNumber = [Basic Product Construction].[New O"
        "racle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8d9403cb205f0c41b969748a5843c42e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.PrimaryUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.Item_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.MaterialQuantity_PerPrimaryUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.MaterialQuanitty_PerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.QTY_UOM"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.OrderNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGetWhereUsedMaterial.Quantity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1153
    Bottom =722
    Left =-1
    Top =-1
    Right =1077
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =266
        Top =73
        Right =617
        Bottom =292
        Top =0
        Name ="tblGetWhereUsedMaterial"
        Name =""
    End
    Begin
        Left =703
        Top =26
        Right =1017
        Bottom =320
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
