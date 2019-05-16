Operation =1
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Expression ="tblBOMOpSequencesInterface.OrganizationCode"
    Expression ="tblBOMOpSequencesInterface.AssemblyItemNumber"
    Expression ="tblBOMOpSequencesInterface.OperationSeqNum"
    Expression ="tblBOMOpSequencesInterface.AlternateRoutingDesignator"
    Expression ="tblBOMOpSequencesInterface.OperationCode"
    Expression ="tblBOMOpSequencesInterface.DepartmentCode"
    Expression ="tblBOMOpSequencesInterface.EffectivityDate"
    Expression ="tblBOMOpSequencesInterface.OperationDescription"
    Expression ="tblBOMOpSequencesInterface.StartupScrap"
    Expression ="tblBOMOpSequencesInterface.ShortText"
    Expression ="tblBOMOpSequencesInterface.CopyCount"
    Expression ="tblBOMOpSequencesInterface.ItemCopied"
    Expression ="tblBOMOpSequencesInterface.OrgOpSeq"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="tblBOMOpSequencesInterface"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber = tblBOMOpSequencesInterface.Op"
        "erationSeqNum"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x340221907bbac8449ed0b03fa42d2297
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.AlternateRoutingDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OperationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.DepartmentCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.StartupScrap"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.ShortText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.EffectivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OperationSeqNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.CopyCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OperationDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OrgOpSeq"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1152
    Bottom =720
    Left =-1
    Top =-1
    Right =1077
    Bottom =554
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =341
        Bottom =215
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =398
        Top =16
        Right =743
        Bottom =220
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End
