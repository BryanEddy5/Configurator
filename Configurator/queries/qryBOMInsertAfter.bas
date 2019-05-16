Operation =3
Name ="tblBOMInventoryCompsInterface"
Option =0
Where ="(((tblBOMInventoryCompsInterface.ItemSeqNumber)=[Forms]![frmItemInterface]![frmI"
    "temInterfaceBOM].[Form]![ItemSeqNumber]))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="OrganizationCode"
    Expression ="tblBOMInventoryCompsInterface.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
    Name ="OperationSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Name ="EffectivityDate"
    Expression ="tblBOMInventoryCompsInterface.EffectivityDate"
    Alias ="ComponentQuantity"
    Name ="ComponentQuantity"
    Expression ="1"
    Name ="AlternateBOMDesignator"
    Expression ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
    Name ="CountPerUOM"
    Expression ="tblBOMInventoryCompsInterface.CountPerUOM"
    Alias ="ItemSeqNumberNew"
    Name ="ItemSeqNumber"
    Expression ="IIf(Forms!frmBOMAfter!NextItemSeqNumber=0,[ItemSeqNumber]+10,[ItemSeqNumber]+(Fo"
        "rms!frmBOMAfter!NextItemSeqNumber-[ItemSeqNumber])/2)"
    Alias ="ComponentItemNumber"
    Name ="ComponentItemNumber"
    Expression ="0"
    Name ="OrgOpSeq"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Name ="Basis"
    Expression ="tblBOMInventoryCompsInterface.Basis"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x34f3ec7cfc7140468b2247906c6b64dc
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentQuantity"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x262d6e7fc45f0b4e8abcd49a5861b66a
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe388844a2dd8534b981cd5034ad90a64
        End
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7c9ab1e43b54849b88576a17a0da0c7
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.Basis"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-712
    Top =113
    Right =795
    Bottom =907
    Left =-1
    Top =-1
    Right =1475
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =282
        Bottom =225
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End
