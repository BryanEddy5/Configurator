Operation =3
Name ="tblBOMInventoryCompsInterface"
Option =0
Begin InputTables
    Name ="tblItemLoadTempBOM"
End
Begin OutputColumns
    Name ="OperationSeqNumber"
    Expression ="tblItemLoadTempBOM.OperationSeqNumber"
    Name ="ItemSeqNumber"
    Expression ="tblItemLoadTempBOM.ItemSeqNumber"
    Name ="OrganizationCode"
    Expression ="tblItemLoadTempBOM.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblItemLoadTempBOM.AssemblyItemNumber"
    Name ="ComponentItemNumber"
    Expression ="tblItemLoadTempBOM.ComponentItemNumber"
    Name ="EffectivityDate"
    Expression ="tblItemLoadTempBOM.EffectivityDate"
    Name ="ComponentQuantity"
    Expression ="tblItemLoadTempBOM.ComponentQuantity"
    Name ="AlternateBOMDesignator"
    Expression ="tblItemLoadTempBOM.AlternateBOMDesignator"
    Name ="CountPerUOM"
    Expression ="tblItemLoadTempBOM.CountPerUOM"
    Name ="LayerID"
    Expression ="tblItemLoadTempBOM.Layer"
    Name ="UnitID"
    Expression ="tblItemLoadTempBOM.UnitID"
    Name ="OrgOpSeq"
    Expression ="tblItemLoadTempBOM.OperationSeqNumber"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="SupplySubinventory"
    Expression ="tblItemLoadTempBOM.SupplySubinventory"
    Name ="SupplyLocator"
    Expression ="tblItemLoadTempBOM.SupplyLocator"
    Name ="SupplyType"
    Expression ="tblItemLoadTempBOM.SupplyType"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![CopyCount]"
    Flag =4
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
End
Begin OrderBy
    Expression ="tblItemLoadTempBOM.ItemSeqNumber"
    Flag =0
    Expression ="tblItemLoadTempBOM.OperationSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb507aa750be2db488d1a436527fa1a21
End
Begin
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x11c66bbc434c6045b1fe1a1cb5da9f15
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe47fca7140858f46af27443b22fb56e2
        End
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ViewOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.SupplyLocator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.SupplySubinventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.SupplyType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =101
    Top =72
    Right =1393
    Bottom =804
    Left =-1
    Top =-1
    Right =1260
    Bottom =301
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =317
        Bottom =259
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
End
