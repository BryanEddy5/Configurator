Operation =3
Name ="tblBOMInventoryCompsInterface"
Option =24
RowCount ="1"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="OrganizationCode"
    Expression ="tblBOMInventoryCompsInterface.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
    Alias ="ColorChip"
    Name ="ComponentItemNumber"
    Expression ="[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![ColorChip]"
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
    Name ="ItemSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
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
Begin OrderBy
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa1207d5fa079c34b996693cc363e20da
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrgOpSeq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.LayerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.EffectivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CopyCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
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
        dbText "Name" ="ColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="CountPerUOM1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f18b986b1478c46a64ff2b44457f67f
        End
    End
    Begin
        dbText "Name" ="ComponentQuantity1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x77b0c19754ba5a4f83c9b028a01c3e03
        End
    End
    Begin
        dbText "Name" ="NewItemSeqNum"
        dbLong "AggregateType" ="-1"
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
    Bottom =349
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
End
