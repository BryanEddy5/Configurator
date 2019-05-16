Operation =3
Name ="tblQAItemSpecElements"
Option =0
Where ="(((tblQASpecsFiberDefaults.Cu_Unit)<=[NumCopperUnits]))"
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
    Name ="tblCableConstructionReferences"
    Name ="tblQAItemSpecs"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="SpecName"
    Expression ="tblQAItemSpecs.SpecName"
    Name ="ElementName"
    Expression ="tblQASpecsFiberDefaults.SpecificationElement"
    Name ="UOMCode"
    Expression ="tblQASpecsFiberDefaults.UOM"
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
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblQAItemSpecs"
    RightTable ="Basic Product Construction"
    Expression ="tblQAItemSpecs.Item = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa71b010ef3d57c44b17ccd5c7bc88327
End
Begin
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.Cu_Unit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultTarget"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.FiberType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumCopperUnits"
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecs.Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUserUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecs.SpecName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =50
    Top =73
    Right =1342
    Bottom =662
    Left =-1
    Top =-1
    Right =1260
    Bottom =209
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =911
        Top =23
        Right =1292
        Bottom =247
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
    Begin
        Left =529
        Top =24
        Right =846
        Bottom =236
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =72
        Top =19
        Right =263
        Bottom =217
        Top =0
        Name ="tblQAItemSpecs"
        Name =""
    End
    Begin
        Left =315
        Top =17
        Right =459
        Bottom =172
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
