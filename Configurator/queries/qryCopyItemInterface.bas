Operation =3
Name ="tblItemInterface"
Option =0
Where ="(((AFLPRD_INVSysItem_CAB.ItemNumber)=Forms!frmMainMenu!ItemNumberSelect) And ((t"
    "blItemInterfaceDataTypes.DataTypeName)=\"Template\") And ((tblItemInterfaceDataT"
    "ypes.DesignType)=10) And ((tblItemInterfaceDataTypes.ItemType)=Forms!frmMainMenu"
    "!NewItemType))"
Begin InputTables
    Name ="tblItemInterfaceDataTypes"
    Name ="vItemSearch"
    Name ="AFLPRD_INVSysItem_CAB"
End
Begin OutputColumns
    Name ="OrganizationCode"
    Expression ="tblItemInterfaceDataTypes.DataTypeOrg"
    Alias ="NewItemNumber"
    Name ="ItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="TemplateName"
    Expression ="tblItemInterfaceDataTypes.DataTypeValue"
    Alias ="PrimaryUOMCode"
    Name ="PrimaryUOMCode"
    Expression ="IIf([Forms]![frmMainMenu]![ItemNumberSelect] Like \"*BX*\",\"EA\",\"M\")"
    Alias ="StdLotSize"
    Name ="StdLotSize"
    Expression ="Forms!frmMainMenu!StdLotSizeNew"
    Alias ="Description"
    Name ="Description"
    Expression ="\".\""
    Alias ="CableWeight"
    Name ="CableWeight"
    Expression ="Nz(BOM_WeightCalculator([ItemNumber]),0)/1000"
    Name ="CableWidth"
    Expression ="vItemSearch.NominalOD"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
End
Begin Joins
    LeftTable ="vItemSearch"
    RightTable ="AFLPRD_INVSysItem_CAB"
    Expression ="vItemSearch.Oracle = AFLPRD_INVSysItem_CAB.ItemNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xacee367277b3c541b5c42c64ea123c85
End
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdb08d172aaf1594da1dd3721e92f3840
        End
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StdLotSize"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5ddc9a56440704491d2684a2bfc2235
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed182b00e4f671489a877238bdb4e4ba
        End
    End
    Begin
        dbText "Name" ="PrimaryUOMCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbffe1a826e847047821c9638387dcc57
        End
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x874baf00a4a49340999731bbda25e897
        End
    End
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.DataTypeOrg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.DataTypeValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItem_CAB.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.ItemType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.DesignType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vItemSearch.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.DataTypeName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vItemSearch.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableWeight"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd517e14926dbb64abeb12811c5222cf3
        End
    End
End
Begin
    State =0
    Left =-18
    Top =59
    Right =1481
    Bottom =896
    Left =-1
    Top =-1
    Right =1467
    Bottom =468
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =358
        Top =7
        Right =633
        Bottom =275
        Top =0
        Name ="tblItemInterfaceDataTypes"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vItemSearch"
        Name =""
    End
    Begin
        Left =681
        Top =12
        Right =987
        Bottom =199
        Top =0
        Name ="AFLPRD_INVSysItem_CAB"
        Name =""
    End
End
