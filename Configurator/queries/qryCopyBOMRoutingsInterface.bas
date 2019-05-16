Operation =3
Name ="tblBOMRoutingsInterface"
Option =0
Where ="(((dbo_AFLPRD_BOMRoutSub_CAB.AssemblyItemNumber)=[Forms]![frmMainMenu]![ItemNumb"
    "erSelect]))"
Begin InputTables
    Name ="qryItemInterfaceLocator"
    Name ="qryItemInterfaceSubinventory"
    Name ="dbo_AFLPRD_BOMRoutSub_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="AssemblyItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="dbo_AFLPRD_BOMRoutSub_CAB.OrganizationCode"
    Name ="CommonItemNumber"
    Expression ="dbo_AFLPRD_BOMRoutSub_CAB.CommonItemNumber"
    Name ="AlternateRoutingDesignator"
    Expression ="dbo_AFLPRD_BOMRoutSub_CAB.AlternateRoutingDesignator"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="CompletionSubinventory"
    Expression ="qryItemInterfaceSubinventory.Subinventory"
    Name ="LocationName"
    Expression ="qryItemInterfaceLocator.LocationName"
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
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x90eb08ad360a35488ed62c2c4955c012
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xffb14e705fc5a64ebefe37cf8c5b78af
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7cdbb2c270a508438cd11d85399bce85
        End
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsInterface.CommonItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMRoutSub_CAB.AlternateRoutingDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMRoutingsInterface.AlternateRoutingDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMRoutSub_CAB.CommonItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemInterfaceSubinventory.Subinventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemInterfaceLocator.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_AFLPRD_BOMRoutSub_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =105
    Top =127
    Right =1418
    Bottom =614
    Left =-1
    Top =-1
    Right =1281
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =789
        Top =22
        Right =1056
        Bottom =182
        Top =0
        Name ="qryItemInterfaceLocator"
        Name =""
    End
    Begin
        Left =449
        Top =27
        Right =729
        Bottom =178
        Top =0
        Name ="qryItemInterfaceSubinventory"
        Name =""
    End
    Begin
        Left =189
        Top =72
        Right =333
        Bottom =216
        Top =0
        Name ="dbo_AFLPRD_BOMRoutSub_CAB"
        Name =""
    End
End
