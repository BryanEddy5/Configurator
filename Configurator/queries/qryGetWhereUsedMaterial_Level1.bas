Operation =3
Name ="tblGetWhereUsedMaterial"
Option =0
Begin InputTables
    Name ="dbo_AFLPRD_BOMInvComp_CAB"
    Name ="tblMaterialBOMExplode"
End
Begin OutputColumns
    Name ="AssemblyItemNumber"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemNumber"
    Name ="Material"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
    Name ="PrimaryUOM"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.PrimaryUOM"
    Name ="MaterialQuantity_PerPrimaryUOM"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentQuantity"
    Name ="CountPerUOM"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.CountPerUOM"
    Name ="Item_Status"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemStatus"
End
Begin Joins
    LeftTable ="tblMaterialBOMExplode"
    RightTable ="dbo_AFLPRD_BOMInvComp_CAB"
    Expression ="tblMaterialBOMExplode.OracleItem = dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
    Flag =1
End
Begin Groups
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemNumber"
    GroupLevel =0
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
    GroupLevel =0
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.PrimaryUOM"
    GroupLevel =0
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentQuantity"
    GroupLevel =0
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.CountPerUOM"
    GroupLevel =0
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemStatus"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x884b5d88facf744fb5386f5edeabb376
End
Begin
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.PrimaryUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =7
    Top =155
    Right =1569
    Bottom =908
    Left =-1
    Top =-1
    Right =1530
    Bottom =325
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =573
        Top =56
        Right =1016
        Bottom =311
        Top =0
        Name ="dbo_AFLPRD_BOMInvComp_CAB"
        Name =""
    End
    Begin
        Left =312
        Top =89
        Right =503
        Bottom =233
        Top =0
        Name ="tblMaterialBOMExplode"
        Name =""
    End
End
