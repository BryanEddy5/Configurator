Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.FiberType) Like [Forms]![frmItemInterface]![frmPremis"
    "eItemDetails].[Form]![FiberType2] & \"*\"))"
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
End
Begin OutputColumns
    Name ="Attribute1"
    Expression ="tblQASpecsFiberDefaults.FiberType"
    Alias ="Fiber2"
    Name ="ElementName"
    Expression ="Replace([SpecificationElement],\"FIBER1\",\"FIBER2\")"
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
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2157b3e74bfb4940b1754c6d1b34da57
End
Begin
    Begin
        dbText "Name" ="Fiber2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x520b4c0d3398864aabc55bdf0acff66e
        End
    End
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x174ab378e3a60a4b995f63cfb471c306
        End
    End
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
    Bottom =214
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
