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
    Expression ="[OperationSeqNum]-([OperationSeqNum]-Forms!frmRoutingBefore!PreviousOpSeqNumber)"
        "/2"
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
    Expression ="IIf(Forms!frmRoutingBefore!PreviousOpSeqNumber=0,[OperationSeqNum]+10,[Operation"
        "SeqNum]-(Forms!frmRoutingBefore!PreviousOpSeqNumber-[OperationSeqNum])/2)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x7d5cf0cf252d9148abb1bf3328274cf8
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpSeqNumberNew"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9535d8b2a1fbf84da16832448be393b2
        End
    End
    Begin
        dbText "Name" ="ResourceSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3768467c2384b24d9d773bbfe3bd048b
        End
    End
    Begin
        dbText "Name" ="ResourceCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ad7530442f6b14795ae38b200beafca
        End
    End
    Begin
        dbText "Name" ="AssignedUnits"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9cf3519286f984e8590125a650fc713
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ebf2606ddc41145a03a520b64934aea
        End
    End
End
Begin
    State =0
    Left =76
    Top =72
    Right =1185
    Bottom =504
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
        Right =306
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End
