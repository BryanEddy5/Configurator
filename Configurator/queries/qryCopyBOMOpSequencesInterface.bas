Operation =3
Name ="tblBOMOpSequencesInterface"
Option =0
Where ="(((dbo_AFLPRD_BOMOpSeq_CAB.AssemblyItemNumber)=[Forms]![frmMainMenu]![ItemNumber"
    "Select]))"
Begin InputTables
    Name ="dbo_AFLPRD_BOMOpSeq_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.OrganizationCode"
    Name ="OperationSeqNum"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.OperationSeqNum"
    Name ="AlternateRoutingDesignator"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.AlternateRoutingDesignator"
    Name ="OperationCode"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.OperationCode"
    Name ="DepartmentCode"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.DepartmentCode"
    Name ="OperationDescription"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.OperationDescription"
    Name ="StartupScrap"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.StartupScrap"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="OrgOpSeq"
    Expression ="dbo_AFLPRD_BOMOpSeq_CAB.OperationSeqNum"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![CopyCount]"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb3e4b466de0794408fabec5db0f3e91b
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x88fc65439d94aa45b29d0023fa4e9750
        End
    End
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd821caeb546e9043846e00f9e42fc0bc
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc0df090dd6963e48b80b1307b75b32ce
        End
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.AlternateRoutingDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.DepartmentCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.StartupScrap"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.OperationSeqNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.OperationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.OperationDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.OperationSeqNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.AlternateRoutingDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.DepartmentCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.StartupScrap"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.OperationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.OperationDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpSequencesInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMOpSeq_CAB.[StartupScrap]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =43
    Right =1418
    Bottom =530
    Left =-1
    Top =-1
    Right =1339
    Bottom =171
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="dbo_AFLPRD_BOMOpSeq_CAB"
        Name =""
    End
End
