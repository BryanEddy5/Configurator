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
    Expression ="[OperationSeqNum]-([OperationSeqNum]-Forms!frmRoutingBefore!PreviousOpSeqNumber)"
        "/2"
    Name ="OrganizationCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode"
    Name ="AssemblyItemNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber"
    Name ="AlternateRoutingDesignator"
    Expression ="tblBOMOpSequencesInterfaceTemp.AlternateRoutingDesignator"
    Name ="StartupScrap"
    Expression ="tblBOMOpSequencesInterfaceTemp.StartupScrap"
    Name ="ShortText"
    Expression ="tblBOMOpSequencesInterfaceTemp.ShortText"
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
    0xfc6b988502110840ba4df032ae589352
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
            0xf2156103bf5a3149b7598c65c0082168
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c1bc00c6b26f74e9e75960ae72fb554
        End
    End
End
Begin
    State =0
    Left =39
    Top =185
    Right =1148
    Bottom =617
    Left =-1
    Top =-1
    Right =1077
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =297
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End
