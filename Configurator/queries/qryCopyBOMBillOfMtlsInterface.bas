Operation =3
Name ="tblBOMBillOfMtlsInterface"
Option =0
Where ="(((AFLPRD_BOMBillMtl_CAB.Assembly)=[Forms]![frmMainMenu]![ItemNumberSelect]))"
Begin InputTables
    Name ="AFLPRD_BOMBillMtl_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="ItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="AFLPRD_BOMBillMtl_CAB.OrganizationCode"
    Name ="AssemblyType"
    Expression ="AFLPRD_BOMBillMtl_CAB.AssemblyType"
    Name ="AlternateBOMDesignator"
    Expression ="AFLPRD_BOMBillMtl_CAB.AlternateBOMDesignator"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xba79dfa854399343952560fce5342e82
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x60c0ce377d36e84dbeee12eaa9eb87fe
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x64c73dc13b2390449850ddef3558a634
        End
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMBillMtl_CAB.Assembly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMBillMtl_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMBillMtl_CAB.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.AssemblyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMBillMtl_CAB.AssemblyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_BOMBillMtl_CAB.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =5
    Right =1302
    Bottom =492
    Left =-1
    Top =-1
    Right =1252
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =345
        Bottom =212
        Top =0
        Name ="AFLPRD_BOMBillMtl_CAB"
        Name =""
    End
End
