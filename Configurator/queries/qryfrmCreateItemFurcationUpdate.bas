Operation =4
Option =0
Where ="(((Mid([Item No],11,1))=[FurcationLetter]))"
Begin InputTables
    Name ="tblNewPrductConstruction"
    Name ="tblCableConstructionReferences"
    Name ="tblFurcationTubeReference"
    Name ="tblDesignCodeFurcationTube"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[TB Material]"
    Expression ="[tblDesignCodeFurcationTube].[Furcation]"
    Name ="tblNewPrductConstruction.[Special Instr Product1]"
    Expression ="[SpecialnInstr]"
    Name ="tblNewPrductConstruction.[TB Nominal OD]"
    Expression ="[tblDesignCodeFurcationTube].[TBNomOD]"
    Name ="tblNewPrductConstruction.[Fiber Oracle item]"
    Expression ="\"NONE\""
End
Begin Joins
    LeftTable ="tblNewPrductConstruction"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblNewPrductConstruction.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblFurcationTubeReference"
    Expression ="tblCableConstructionReferences.BaseID = tblFurcationTubeReference.BaseID"
    Flag =1
    LeftTable ="tblFurcationTubeReference"
    RightTable ="tblDesignCodeFurcationTube"
    Expression ="tblFurcationTubeReference.Furcation = tblDesignCodeFurcationTube.Furcation"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xfc3c8c64fbf7724184cb0e5018e52fc1
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Material]"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Special Instr Product1]"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Nominal OD]"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.Furcation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mid([Item No],11,1)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.SpecialnInstr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =132
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =796
        Top =13
        Right =1238
        Bottom =206
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
    Begin
        Left =553
        Top =19
        Right =709
        Bottom =194
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =363
        Top =31
        Right =507
        Bottom =175
        Top =0
        Name ="tblFurcationTubeReference"
        Name =""
    End
    Begin
        Left =110
        Top =25
        Right =332
        Bottom =173
        Top =0
        Name ="tblDesignCodeFurcationTube"
        Name =""
    End
End
