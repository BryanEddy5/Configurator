dbMemo "SQL" ="INSERT INTO tblGetWhereUsedMaterial ( Material, AssemblyItemNumber, MaterialQuan"
    "tity_PerPrimaryUOM, MaterialQuanitty_PerOrder, PrimaryUOM, Item_Status, Customer"
    ", OrderNum, Shipped, Quantity, QTY_UOM, REVENUE, Description, SoLineNumber, Tran"
    "sactionDate )\015\012SELECT qryGetBuyMaterials.Material, qryGetBuyMaterials.Asse"
    "mblyItemNumber, qryGetBuyMaterials.MaterialQuantity_PerPrimaryUOM, qryGetBuyMate"
    "rials.MaterialQuanitty_PerOrder, qryGetBuyMaterials.PrimaryUOM, qryGetBuyMateria"
    "ls.Item_Status, qryGetBuyMaterials.Customer, qryGetBuyMaterials.OrderNum, qryGet"
    "BuyMaterials.Shipped, qryGetBuyMaterials.Quantity, qryGetBuyMaterials.QTY_UOM, q"
    "ryGetBuyMaterials.REVENUE, qryGetBuyMaterials.Description, qryGetBuyMaterials.SO"
    "_LINE_NUMBER, qryGetBuyMaterials.TRANSACTION_DATE\015\012FROM qryGetBuyMaterials"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0x37e2c9fa8b1a184e99ffa042cca312fb
End
Begin
    Begin
        dbText "Name" ="qryGetBuyMaterials.StartUpScrap"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.OracleItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.QTY_UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.OrderNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.REVENUE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.MaterialQuantity_PerPrimaryUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.PrimaryUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Item_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.MaterialQuanitty_PerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.SO_LINE_NUMBER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.TRANSACTION_DATE"
        dbLong "AggregateType" ="-1"
    End
End
