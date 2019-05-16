Operation =3
Name ="tblQAItemSpecElements"
Option =2
Where ="(((AFLPRD_INVSysItemSpec_CAB.SpecificationElement) Not Like \"fiber*\" And (AFLP"
    "RD_INVSysItemSpec_CAB.SpecificationElement) Not Like \"*ET_COND*\" And (AFLPRD_I"
    "NVSysItemSpec_CAB.SpecificationElement) Not Like \"*efl*\") And ((AFLPRD_INVSysI"
    "temSpec_CAB.ItemNumber)=Forms!frmMainMenu!ItemNumberSelect))"
Begin InputTables
    Name ="AFLPRD_INVSysItemSpec_CAB"
End
Begin OutputColumns
    Alias ="SpecName"
    Name ="SpecName"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Alias ="OrganizationCode"
    Name ="OrganizationCode"
    Expression ="\"CAB\""
    Name ="ElementName"
    Expression ="AFLPRD_INVSysItemSpec_CAB.SpecificationElement"
    Name ="UOMCode"
    Expression ="AFLPRD_INVSysItemSpec_CAB.UOM"
    Name ="TargetValue"
    Expression ="AFLPRD_INVSysItemSpec_CAB.TargetValue"
    Name ="LowerSpecLimit"
    Expression ="AFLPRD_INVSysItemSpec_CAB.LowerLimit"
    Name ="UpperSpecLimit"
    Expression ="AFLPRD_INVSysItemSpec_CAB.UpperLimit"
    Name ="LowerUserDefinedLimit"
    Expression ="AFLPRD_INVSysItemSpec_CAB.UserLowerLimit"
    Name ="UpperUserDefinedLimit"
    Expression ="AFLPRD_INVSysItemSpec_CAB.UserUpperLimit"
    Name ="Attribute1"
    Expression ="AFLPRD_INVSysItemSpec_CAB.FiberType"
    Name ="Attribute2"
    Expression ="AFLPRD_INVSysItemSpec_CAB.TestLength"
    Name ="Attribute3"
    Expression ="AFLPRD_INVSysItemSpec_CAB.TestIncrement"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb71d197333f9254db7cde70aec5d2114
End
Begin
    Begin
        dbText "Name" ="OrganizationCode"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee7292849c336c40b5acc03d5a668c02
        End
    End
    Begin
        dbText "Name" ="SpecName"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x57623c4d0169c740ada3aa527484dd04
        End
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.TargetValue"
        dbInteger "ColumnWidth" ="5895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.TargetValue"
        dbInteger "ColumnWidth" ="5895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.UserUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.FiberType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.LowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.UpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.UserLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.TestLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.TestIncrement"
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
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =534
        Bottom =244
        Top =0
        Name ="AFLPRD_INVSysItemSpec_CAB"
        Name =""
    End
End
