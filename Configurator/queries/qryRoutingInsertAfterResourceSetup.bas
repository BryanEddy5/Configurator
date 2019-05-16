Operation =3
Name ="tblBOMOpResourcesInterface"
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OperationSeqNum)=[Forms]![frmItemInterface]![f"
    "rmItemInterfaceRouting].[Form]![OperationSeqNumber]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Alias ="OpSeqNumberNew"
    Name ="OperationSeqNumber"
    Expression ="IIf(Forms!frmRoutingAfter!NextOpSeqNumber=0,[OperationSeqNum]+10,[OperationSeqNu"
        "m]+(Forms!frmRoutingAfter!NextOpSeqNumber-[OperationSeqNum])/2)"
    Name ="OrganizationCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber"
    Name ="AlternateRoutingDesignator"
    Expression ="tblBOMOpSequencesInterfaceTemp.AlternateRoutingDesignator"
    Alias ="ResourceSeqNumber"
    Name ="ResourceSeqNumber"
    Expression ="2"
    Alias ="ResourceCode"
    Name ="ResourceCode"
    Expression ="\"SETUP\""
    Alias ="AssignedUnits"
    Name ="AssignedUnits"
    Expression ="1"
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
    0x851e6521b91bca4ea9614c4abe9cc44c
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
            0x6497edca845d7741ac942302b883d128
        End
    End
    Begin
        dbText "Name" ="ResourceSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x78bce4dff014a74283e4eb371229fc0a
        End
    End
    Begin
        dbText "Name" ="ResourceCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x63fe8e461c28664097a929191e3c1bbf
        End
    End
    Begin
        dbText "Name" ="AssignedUnits"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c82e831728a5d4da55e6d1d50e68884
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3b11679e57691a468e6771957a95b006
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
        Right =309
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End
