Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.FiberType) Like [Forms]![frmItemInterface]![frmPremis"
    "eItemDetails].[Form]![FiberType1] & \"*\"))"
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
    Name ="Attribute2"
    Expression ="tblQASpecsFiberDefaults.TestLength"
    Name ="Attribute3"
    Expression ="tblQASpecsFiberDefaults.TestIncrement"
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
    0x5c3644379da28449b68e942fd9a528d5
End
Begin
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.FiberType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.TestLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.TestIncrement"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x57783d3ba4bd814b91c75d62dd787041
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =231
    Left =1000
    Top =0
    ColumnsShown =651
    Begin
        Left =-1553
        Top =6
        Right =-1319
        Bottom =266
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
End
