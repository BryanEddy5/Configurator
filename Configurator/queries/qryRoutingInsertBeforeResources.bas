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
    Expression ="1"
    Alias ="ResourceCode"
    Name ="ResourceCode"
    Expression ="0"
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
    0xac2e5d6d7387a84c8df80aa3d5d18831
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
            0x0c37e7d5e65b984b97d21cd6e95d9bd2
        End
    End
    Begin
        dbText "Name" ="ResourceSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4ca761030f032b43bc865f2795ec7851
        End
    End
    Begin
        dbText "Name" ="ResourceCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd50972251eb5948bab7885769e6e9ed
        End
    End
    Begin
        dbText "Name" ="AssignedUnits"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb443ef9f274d642a583d7133ba03345
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x84f1862788689c418673bc3215348c8f
        End
    End
End
Begin
    State =0
    Left =134
    Top =130
    Right =1243
    Bottom =562
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
