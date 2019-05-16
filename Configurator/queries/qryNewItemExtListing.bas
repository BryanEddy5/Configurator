Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.FiberType) Like [Forms]![frmItemInterface]![frmPremis"
    "eItemDetailsPrintExt].[Form]![Listing]))"
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
    0xbc028338314d274ea3261fcd7f590ee4
End
Begin
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3960664ecd20a24f9db9a9d52689c019
        End
    End
End
Begin
    State =0
    Left =237
    Top =31
    Right =1529
    Bottom =493
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
