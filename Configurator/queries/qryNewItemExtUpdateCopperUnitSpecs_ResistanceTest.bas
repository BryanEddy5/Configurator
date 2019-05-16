Operation =4
Option =8
Where ="(((tblQAItemSpecElements.ElementName) Like \"*\" & [SpecificationElement] & \"*\""
    "))"
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
    Name ="Basic Product Construction"
    Name ="tblQAItemSpecElements"
    Name ="Product Desrcriptions"
    Name ="tblItemInterface"
    Name ="tblCableConstructionReferences"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
End
Begin OutputColumns
    Name ="tblQAItemSpecElements.UpperSpecLimit"
    Expression ="[DefaultUpperLimit]"
    Name ="tblQAItemSpecElements.LowerSpecLimit"
    Expression ="[DefaultLowerLimit]"
    Name ="tblQAItemSpecElements.UpperUserDefinedLimit"
    Expression ="[DefaultUserUpperLimit]"
    Name ="tblQAItemSpecElements.LowerUserDefinedLimit"
    Expression ="[DefaultUserLowerLimit]"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblQASpecsFiberDefaults"
    Expression ="[Basic Product Construction].Base = tblQASpecsFiberDefaults.FiberType"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Product Desrcriptions"
    Expression ="[Basic Product Construction].Base = [Product Desrcriptions].[Level 1 Product]"
    Flag =1
    LeftTable ="Product Desrcriptions"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Product Desrcriptions].ProductID = tblCableConstructionReferences.SetupID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Basic Product Construction_1"
    Expression ="tblCableConstructionReferences.Base = [Basic Product Construction_1].Base"
    Flag =1
    LeftTable ="Basic Product Construction_1"
    RightTable ="tblItemInterface"
    Expression ="[Basic Product Construction_1].[New Oracle Part #] = tblItemInterface.ItemNumber"
    Flag =1
    LeftTable ="tblItemInterface"
    RightTable ="tblQAItemSpecElements"
    Expression ="tblItemInterface.ItemNumber = tblQAItemSpecElements.SpecName"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xe47f86185c6fde499c74b1f754e9f1df
End
Begin
    Begin
        dbText "Name" ="tblQAItemSpecElements.ElementName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUserUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.SpecificationElement"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUserLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultUpperLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.UpperSpecLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQASpecsFiberDefaults.DefaultLowerLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.LowerSpecLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.UpperUserDefinedLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.LowerUserDefinedLimit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =55
    Top =76
    Right =1347
    Bottom =808
    Left =-1
    Top =-1
    Right =1260
    Bottom =472
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =981
        Top =192
        Right =1222
        Bottom =417
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
    Begin
        Left =694
        Top =155
        Right =938
        Bottom =466
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =1244
        Top =44
        Right =1454
        Bottom =380
        Top =0
        Name ="tblQAItemSpecElements"
        Name =""
    End
    Begin
        Left =511
        Top =130
        Right =679
        Bottom =458
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =956
        Top =15
        Right =1150
        Bottom =312
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
    Begin
        Left =271
        Top =51
        Right =476
        Bottom =462
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =576
        Top =15
        Right =879
        Bottom =458
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
End
