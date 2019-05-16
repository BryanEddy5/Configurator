Operation =3
Name ="tblBOMOpResourcesInterface"
Option =0
Where ="(((tblNetworkBOMOpResourcesInterface.AssemblyItemNumber)=[Forms]![frmMainMenu]!["
    "ItemNumberSelect]))"
Begin InputTables
    Name ="tblNetworkBOMOpResourcesInterface"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="tblNetworkBOMOpResourcesInterface.OrganizationCode"
    Name ="AlternateRoutingDesignator"
    Expression ="tblNetworkBOMOpResourcesInterface.AlternateRoutingDesignator"
    Name ="OperationSeqNumber"
    Expression ="tblNetworkBOMOpResourcesInterface.OperationSeqNumber"
    Name ="ResourceSeqNumber"
    Expression ="tblNetworkBOMOpResourcesInterface.ResourceSeqNumber"
    Name ="ResourceCode"
    Expression ="tblNetworkBOMOpResourcesInterface.ResourceCode"
    Name ="StandardRateFlag"
    Expression ="tblNetworkBOMOpResourcesInterface.StandardRateFlag"
    Name ="AssignedUnits"
    Expression ="tblNetworkBOMOpResourcesInterface.AssignedUnits"
    Name ="UsageRateOrAmount"
    Expression ="tblNetworkBOMOpResourcesInterface.UsageRateOrAmount"
    Name ="UsageRateOrAmountInverse"
    Expression ="tblNetworkBOMOpResourcesInterface.UsageRateOrAmountInverse"
    Name ="BasisType"
    Expression ="tblNetworkBOMOpResourcesInterface.BasisType"
    Name ="ScheduleFlag"
    Expression ="tblNetworkBOMOpResourcesInterface.ScheduleFlag"
    Name ="AutochargeType"
    Expression ="tblNetworkBOMOpResourcesInterface.AutochargeType"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="OrgOpSeq"
    Expression ="tblNetworkBOMOpResourcesInterface.OperationSeqNumber"
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
dbBinary "GUID" = Begin
    0x675fa69e1436e04aaaa1e6bc6bf6322d
End
Begin
    Begin
        dbText "Name" ="tblOracleBOMOpResourcesInterface.ResourceSeqNumber"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe86b9d7f7e1a094e9a9ab76c35fbfa36
        End
    End
    Begin
        dbText "Name" ="TBLOracleRTGRESOURCES.ResourceSeqNumber"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBLOracleRTGRESOURCES.RESOURCE_SEQ_NUM"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMOpResourcesInterface.ResourceSeqNumber"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3dd7fc81a243284d9432224ed4c362cc
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xffdd96efce07234d8358258ba01842ff
        End
    End
End
Begin
    State =0
    Left =76
    Top =72
    Right =1418
    Bottom =559
    Left =-1
    Top =-1
    Right =1310
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =332
        Bottom =109
        Top =0
        Name ="tblNetworkBOMOpResourcesInterface"
        Name =""
    End
End
