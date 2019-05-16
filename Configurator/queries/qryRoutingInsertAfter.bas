Operation =3
Name ="tblBOMOpSequencesInterfaceTemp"
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OperationSeqNum)=[Forms]![frmItemInterface]![f"
    "rmItemInterfaceRouting].[Form]![OperationSeqNumber]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Alias ="OpSeqNumberNew"
    Name ="OperationSeqNum"
    Expression ="IIf(Forms!frmRoutingAfter!NextOpSeqNumber=0,[OperationSeqNum]+10,[OperationSeqNu"
        "m]+(Forms!frmRoutingAfter!NextOpSeqNumber-[OperationSeqNum])/2)"
    Name ="OrganizationCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber"
    Name ="AlternateRoutingDesignator"
    Expression ="tblBOMOpSequencesInterfaceTemp.AlternateRoutingDesignator"
    Alias ="OrgOpSeqNew"
    Name ="OrgOpSeq"
    Expression ="IIf(Forms!frmRoutingAfter!NextOpSeqNumber=0,[OperationSeqNum]+10,[OperationSeqNu"
        "m]+(Forms!frmRoutingAfter!NextOpSeqNumber-[OperationSeqNum])/2)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x1e96e9e2eae05b4bb5096dbdc1ceea9e
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
    End
    Begin
        dbText "Name" ="ItemSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfd4c4b7bc44a8f49a7d7c215a4c621ff
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03071539a7d32d4cb639bed9832f61dd
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =222
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =304
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End
