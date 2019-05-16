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
    Expression ="1"
    Alias ="ResourceCode"
    Name ="ResourceCode"
    Expression ="0"
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
    0xc04c4af7e92be744b71afec968e69ffe
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
            0xb840c7dadc53334d879b3fc54e7d0a24
        End
    End
    Begin
        dbText "Name" ="ResourceSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7de35b4902085243a8ce283ee1de72ad
        End
    End
    Begin
        dbText "Name" ="ResourceCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8464eb69752e045a0dc0db078ceaeb6
        End
    End
    Begin
        dbText "Name" ="AssignedUnits"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7212b723fcf47e44b6ac730eb1161aeb
        End
    End
    Begin
        dbText "Name" ="OrgOpSeqNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7abeef676e0b744998003b202a6fb32b
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
        Right =321
        Bottom =225
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End
