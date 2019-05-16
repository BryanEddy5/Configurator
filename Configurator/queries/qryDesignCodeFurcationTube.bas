Operation =1
Option =0
Where ="(((tblCableConstructionReferences.Base)=[Forms]![frmDesignCode]![OD]))"
Begin InputTables
    Name ="tblDesignCodeFurcationTube"
    Name ="tblFurcationTubeReference"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="tblDesignCodeFurcationTube.*"
    Expression ="tblCableConstructionReferences.Base"
End
Begin Joins
    LeftTable ="tblDesignCodeFurcationTube"
    RightTable ="tblFurcationTubeReference"
    Expression ="tblDesignCodeFurcationTube.Furcation = tblFurcationTubeReference.Furcation"
    Flag =1
    LeftTable ="tblFurcationTubeReference"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblFurcationTubeReference.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x637d8aae42e00b4986167fd72ff0c591
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.FurcationLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.FurcationTube"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.Furcation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.TBNomOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.SpecialnInstr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =271
    Top =141
    Right =1065
    Bottom =435
    Left =-1
    Top =-1
    Right =762
    Bottom =196
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblDesignCodeFurcationTube"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tblFurcationTubeReference"
        Name =""
    End
    Begin
        Left =644
        Top =48
        Right =910
        Bottom =526
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
