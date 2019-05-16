Operation =1
Option =8
Begin InputTables
    Name ="dbo_AFLPRD_onhandqty_STD_81_ALL"
    Name ="tblGetWhereUsedMaterial"
End
Begin OutputColumns
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.Item"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.UOM"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.QTY"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Matl Value]"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Sub-inventory]"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Material-Account]"
    Expression ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Org Code]"
End
Begin Joins
    LeftTable ="tblGetWhereUsedMaterial"
    RightTable ="dbo_AFLPRD_onhandqty_STD_81_ALL"
    Expression ="tblGetWhereUsedMaterial.AssemblyItemNumber = dbo_AFLPRD_onhandqty_STD_81_ALL.Ite"
        "m"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe8e29d139dc99640906b26e748823716
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qryGetWhereUsedMaterialOnHnadInventory].[Item], [qryGetWhereUsedMaterialOnHnadI"
    "nventory].[Sub-inventory]"
Begin
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.QTY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Matl Value]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Sub-inventory]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Material-Account]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_onhandqty_STD_81_ALL.[Org Code]"
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
    Bottom =299
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =235
        Top =1
        Right =441
        Bottom =453
        Top =0
        Name ="dbo_AFLPRD_onhandqty_STD_81_ALL"
        Name =""
    End
    Begin
        Left =545
        Top =39
        Right =848
        Bottom =410
        Top =0
        Name ="tblGetWhereUsedMaterial"
        Name =""
    End
End
