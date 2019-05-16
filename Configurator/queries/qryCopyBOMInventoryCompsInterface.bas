Operation =3
Name ="tblBOMInventoryCompsInterface"
Option =0
Where ="(((dbo_AFLPRD_BOMInvComp_CAB.AlternateBOMDesignator) Is Null) And ((dbo_AFLPRD_B"
    "OMInvComp_CAB.AssemblyItemNumber)=Forms!frmMainMenu!ItemNumberSelect))"
Begin InputTables
    Name ="dbo_AFLPRD_BOMInvComp_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.OrganizationCode"
    Name ="ComponentItemNumber"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
    Name ="OperationSeqNumber"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.OperationSeqNumber"
    Name ="ComponentQuantity"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ComponentQuantity"
    Name ="CountPerUOM"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.CountPerUOM"
    Name ="ItemSeqNumber"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ItemSeqNumber"
    Name ="UnitID"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.UnitID"
    Name ="LayerID"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.LayerID"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="OrgOpSeq"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.OperationSeqNumber"
    Alias ="EffectivityDate"
    Name ="EffectivityDate"
    Expression ="Date()"
    Name ="AlternateBOMDesignator"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.AlternateBOMDesignator"
    Alias ="Basis2"
    Name ="Basis"
    Expression ="IIf([Basis]=\"Lot\",\"2\",Null)"
    Name ="SupplyLocator"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.SupplyLocator"
    Name ="SupplySubinventory"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.SupplySubinventory"
    Name ="SupplyType"
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.SupplyType"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![ItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![CopyCount]"
    Flag =4
End
Begin OrderBy
    Expression ="dbo_AFLPRD_BOMInvComp_CAB.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x90820a8b443c5a40bf432d0008e59e00
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf47dc1e654b3254497365611b3b25dbb
        End
    End
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5af2f05b1e6ab046bbed44f859d2c44b
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x20e623d35efa4b4d85fc0dcdedc4ef41
        End
    End
    Begin
        dbText "Name" ="EffectivityDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0b6882fbcb42fe49bd84782521516340
        End
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.LayerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.LayerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.[CountPerUOM]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.EffectivityDateTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.Basis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basis2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2734ed0facd66e4ca7eb90be8f60e8e1
        End
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.SupplyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.SupplySubinventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMInvComp_CAB.SupplyLocator"
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
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =583
        Bottom =211
        Top =0
        Name ="dbo_AFLPRD_BOMInvComp_CAB"
        Name =""
    End
End
