Operation =3
Name ="tblBOMRoutingsAttachments"
Option =0
Where ="(((AFLPRD_BOMRoutAtc_CAB.AssemblyItemNumber)=[Forms]![frmMainMenu]![ItemNumberSe"
    "lect]))"
Begin InputTables
    Name ="AFLPRD_BOMRoutAtc_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="AFLPRD_BOMRoutAtc_CAB.OrganizationCode"
    Name ="OperationSeqNumber"
    Expression ="AFLPRD_BOMRoutAtc_CAB.OperationSeqNumber"
    Name ="SeqNumber"
    Expression ="AFLPRD_BOMRoutAtc_CAB.SeqNumber"
    Name ="ShortText"
    Expression ="AFLPRD_BOMRoutAtc_CAB.ShortText"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="OrgOpSeq"
    Expression ="AFLPRD_BOMRoutAtc_CAB.OperationSeqNumber"
    Name ="Description"
    Expression ="AFLPRD_BOMRoutAtc_CAB.Description"
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
    0xeec0f163e4faa443ab8b1a1157f5bc1c
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde2db1ba4519bf4b845cc006d92ed8af
        End
    End
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x13c48d7248937b40bb0188661140deb7
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19eb6ad5d9760445945fb29582f4301f
        End
    End
    Begin
        dbText "Name" ="AFLPRD_BOMRoutAtc_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMRoutAtc_CAB.SeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMRoutAtc_CAB.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMRoutAtc_CAB.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMRoutAtc_CAB.ShortText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.SeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.ShortText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsAttachments.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =387
        Top =12
        Right =531
        Bottom =156
        Top =0
        Name ="AFLPRD_BOMRoutAtc_CAB"
        Name =""
    End
End
