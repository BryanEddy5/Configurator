Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.FiberType)=[Forms]![frmItemInterfaceQTSFiberSelect]!["
    "cboFiberSelect]))"
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
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBinary "GUID" = Begin
    0xa0ce93f97bf82749a83c59ab2574f404
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
            0xe2fbdf9d00109c4fb517de1fcd212726
        End
    End
End
Begin
    State =0
    Left =72
    Top =105
    Right =1181
    Bottom =734
    Left =-1
    Top =-1
    Right =1077
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
