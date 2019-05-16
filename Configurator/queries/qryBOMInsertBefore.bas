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
    Expression ="[ItemSeqNumber]-([ItemSeqNumber]-Forms!frmBOMBefore!PreviousItemSeqNumber)/2"
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
    0x48af2fd95d8609478f00ec37ca5d0016
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
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x650dee2c878e1c4cacb0a6e81fca932d
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumberNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xacd2000b2cc9834eaf38e68940c26255
        End
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff3f6f378cbda94eab03e7155e4ec3d8
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.Basis"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =163
    Top =185
    Right =1418
    Bottom =779
    Left =-1
    Top =-1
    Right =1223
    Bottom =198
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
