dbMemo "SQL" ="INSERT INTO tblGetBuyMaterials ( OracleItem, Quantity, ComponentItemNumber, Comp"
    "onentQuantity_PerQty, TotalCost, PrimaryUOM, ComponenetQuantity_M, StartUpScrap "
    ")\015\012SELECT qryGetBuyMaterials.OracleItem AS Expr1, qryGetBuyMaterials.Quant"
    "ity, qryGetBuyMaterials.ComponentItemNumber AS Expr2, qryGetBuyMaterials.Compone"
    "ntQuantity_PerQty AS Expr3, qryGetBuyMaterials.TotalCost AS Expr4, qryGetBuyMate"
    "rials.PrimaryUOM, qryGetBuyMaterials.ComponentQuantity_M AS Expr5, qryGetBuyMate"
    "rials.StartUpScrap AS Expr6\015\012FROM qryGetBuyMaterials;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xf64734134b905e49b3f783318169ad25
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
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
        dbBinary "GUID" = Begin
            0x9d1384ab3134fe4d9741a0c17ac6bbd3
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d56b1ca41d4e94580baca88f2bfb3e3
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a665a8c8c98d54181797153bddbfee5
        End
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.ComponentQuantity_M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4d7d18617f38f4b82a78b7f174a16f0
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfeaaa9973184324a86cbef5dfba3497f
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb74b2cac54023947b99f1934fcce30bc
        End
    End
    Begin
        dbText "Name" ="qryGetBuyMaterials.Description"
        dbLong "AggregateType" ="-1"
    End
End
