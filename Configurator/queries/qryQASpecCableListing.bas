Operation =3
Name ="tblQAItemSpecElements"
Option =2
Where ="(((tblQASpecsFiberDefaults.SpecificationElement)=\"CABLE LISTING\"))"
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
    Alias ="CableRating"
    Name ="TargetValue"
    Expression ="CableRatingPrint([Print Line 2] & \" \" & [Print Line 3] & \" \",[Item No])"
End
Begin Joins
    LeftTable ="tblItemInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblItemInterface.ItemNumber = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0x812fedbf5173f441980bb0d4c1dd4cf5
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
    Begin
        dbText "Name" ="[Basic Product Construction].ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRating"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
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
    Bottom =206
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =166
        Top =39
        Right =422
        Bottom =183
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
    Begin
        Left =950
        Top =44
        Right =1416
        Bottom =268
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
    Begin
        Left =591
        Top =30
        Right =869
        Bottom =247
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
