Operation =3
Name ="tblItemCatalogInterface"
Option =0
Where ="(((AFLPRD_INVItmCtlg_CAB.Item)=[Forms]![frmMainMenu]![ItemNumberSelect]))"
Begin InputTables
    Name ="AFLPRD_INVItmCtlg_CAB"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="ItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="ElementName"
    Expression ="AFLPRD_INVItmCtlg_CAB.ElementName"
    Name ="ElementValue"
    Expression ="AFLPRD_INVItmCtlg_CAB.ElementValue"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Alias ="AllowOverride"
    Name ="AllowOverride"
    Expression ="-1"
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
    0xd0333628344f1c48ad0c923c6a4da84f
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9088791e5855e74085acdaad7f9d9d57
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8aedb91b658e5e4aa6c529655c766713
        End
    End
    Begin
        dbText "Name" ="AllowOverride"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4de191c92d03d04c816463262a1d8817
        End
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCtlg_CAB.ElementName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCtlg_CAB.Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemCatalogInterface.ElementName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemCatalogInterface.ElementValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemCatalogInterface.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCtlg_CAB.ElementValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCtlg_CAB.ItemNumber"
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
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="AFLPRD_INVItmCtlg_CAB"
        Name =""
    End
End
