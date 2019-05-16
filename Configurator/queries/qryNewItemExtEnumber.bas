Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.DefaultTarget)=[Forms]![frmItemInterface]![frmPremise"
    "ItemDetailsPrintExt].[Form]![Enumber]))"
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
End
Begin OutputColumns
    Name ="Attribute1"
    Expression ="tblQASpecsFiberDefaults.FiberType"
    Name ="ElementName"
    Expression ="tblQASpecsFiberDefaults.SpecificationElement"
    Name ="TargetValue"
    Expression ="tblQASpecsFiberDefaults.DefaultTarget"
    Name ="UpperSpecLimit"
    Expression ="tblQASpecsFiberDefaults.DefaultUpperLimit"
    Name ="LowerSpecLimit"
    Expression ="tblQASpecsFiberDefaults.DefaultLowerLimit"
    Name ="UpperUserDefinedLimit"
    Expression ="tblQASpecsFiberDefaults.DefaultUserUpperLimit"
    Name ="LowerUserDefinedLimit"
    Expression ="tblQASpecsFiberDefaults.DefaultUserLowerLimit"
    Alias ="SpecName"
    Name ="SpecName"
    Expression ="Forms!frmItemInterface!ItemNumber"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb61400470196b64bbbf859cee719a52c
End
Begin
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3adb22f155f41a41b348e993e21d9fe4
        End
    End
End
Begin
    State =0
    Left =253
    Top =82
    Right =1545
    Bottom =544
    Left =-1
    Top =-1
    Right =1260
    Bottom =231
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =249
        Bottom =169
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
End
