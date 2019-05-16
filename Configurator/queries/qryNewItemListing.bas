Operation =3
Name ="tblQAItemSpecElements"
Option =2
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
    Name ="tblItemInterface"
    Name ="Basic Product Construction"
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
    Name ="SpecName"
    Expression ="tblItemInterface.ItemNumber"
End
Begin Joins
    LeftTable ="tblItemInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblItemInterface.ItemNumber = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
    LeftTable ="tblQASpecsFiberDefaults"
    RightTable ="Basic Product Construction"
    Expression ="tblQASpecsFiberDefaults.FiberType = [Basic Product Construction].[Listing Compan"
        "y]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc36a724df6c8aa4ca7a77b592348d1fe
End
Begin
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3960664ecd20a24f9db9a9d52689c019
        End
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.FiberType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUserUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultTarget"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterface.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.SpecificationElement"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUserLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-553
    Top =41
    Right =556
    Bottom =503
    Left =-1
    Top =-1
    Right =1077
    Bottom =308
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =84
        Top =12
        Right =340
        Bottom =156
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
    Begin
        Left =365
        Top =16
        Right =831
        Bottom =240
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
    Begin
        Left =957
        Top =17
        Right =1235
        Bottom =234
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
