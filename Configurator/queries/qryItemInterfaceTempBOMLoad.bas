Operation =3
Name ="tblItemLoadTempBOM"
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="tblBOMInventoryCompsInterface.OrganizationCode"
    Name ="ComponentItemNumber"
    Expression ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Name ="OperationSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Name ="ComponentQuantity"
    Expression ="tblBOMInventoryCompsInterface.ComponentQuantity"
    Name ="AlternateBOMDesignator"
    Expression ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
    Name ="CountPerUOM"
    Expression ="tblBOMInventoryCompsInterface.CountPerUOM"
    Name ="ItemSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Name ="UnitID"
    Expression ="tblBOMInventoryCompsInterface.UnitID"
    Name ="Layer"
    Expression ="tblBOMInventoryCompsInterface.LayerID"
    Alias ="EffectivityDate"
    Name ="EffectivityDate"
    Expression ="Date()"
    Name ="Basis"
    Expression ="tblBOMInventoryCompsInterface.Basis"
    Name ="SupplySubinventory"
    Expression ="tblBOMInventoryCompsInterface.SupplySubinventory"
    Name ="SupplyLocator"
    Expression ="tblBOMInventoryCompsInterface.SupplyLocator"
    Name ="SupplyType"
    Expression ="tblBOMInventoryCompsInterface.SupplyType"
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
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Flag =0
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd34bf41cb3396e4e8dfe8a139784ac86
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x74d3f605f893b046887b06fa86317a7d
        End
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ItemSeqNumber"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EffectivityDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe73ecc5a6dc240459fab18cd0310378e
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.Basis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
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
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplySubinventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplyLocator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.SupplyType"
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
End
Begin
    State =0
    Left =49
    Top =95
    Right =1543
    Bottom =855
    Left =-1
    Top =-1
    Right =1462
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =521
        Bottom =270
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End
