Operation =3
Name ="tblBOMInventoryCompsInterface"
Option =24
RowCount ="1"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="OrganizationCode"
    Expression ="tblBOMInventoryCompsInterface.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
    Alias ="FurcationTube"
    Name ="ComponentItemNumber"
    Expression ="[Basic Product Construction].[TB Material]"
    Name ="OperationSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Name ="EffectivityDate"
    Expression ="tblBOMInventoryCompsInterface.EffectivityDate"
    Alias ="ComponentQuantity1"
    Name ="ComponentQuantity"
    Expression ="0"
    Name ="AlternateBOMDesignator"
    Expression ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
    Alias ="CountPerUOM1"
    Name ="CountPerUOM"
    Expression ="1"
    Alias ="ItemSeqNumbers"
    Name ="ItemSeqNumber"
    Expression ="10"
    Name ="UnitID"
    Expression ="tblBOMInventoryCompsInterface.UnitID"
    Name ="LayerID"
    Expression ="tblBOMInventoryCompsInterface.LayerID"
    Name ="CopyCount"
    Expression ="tblBOMInventoryCompsInterface.CopyCount"
    Name ="ItemCopied"
    Expression ="tblBOMInventoryCompsInterface.ItemCopied"
    Name ="OrgOpSeq"
    Expression ="tblBOMInventoryCompsInterface.OrgOpSeq"
    Name ="Basis"
    Expression ="tblBOMInventoryCompsInterface.Basis"
    Name ="SupplySubinventory"
    Expression ="tblBOMInventoryCompsInterface.SupplySubinventory"
    Name ="SupplyLocator"
    Expression ="tblBOMInventoryCompsInterface.SupplyLocator"
    Name ="SupplyType"
    Expression ="tblBOMInventoryCompsInterface.SupplyType"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber = [Basic Product Construction]."
        "[New Oracle Part #]"
    Flag =1
End
Begin OrderBy
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x549774625e3afb4aa549619c1815238f
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrgOpSeq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.EffectivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewItemSeqNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CopyCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.LayerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountPerUOM1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc2d4977318668b4c95edf0c03a15f6bc
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewItemSeq"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="ComponentQuantity1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5e6663282b2ca140ba09ca78c6c4dfec
        End
    End
    Begin
        dbText "Name" ="ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FurcationTube"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumbers"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6a60762c5dbbfc4dbf7b39789301b595
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.Basis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplyLocator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplySubinventory"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1577
    Bottom =965
    Left =-1
    Top =-1
    Right =1553
    Bottom =416
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =449
        Bottom =343
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =497
        Top =12
        Right =829
        Bottom =275
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
