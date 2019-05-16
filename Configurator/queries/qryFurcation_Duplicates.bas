Operation =1
Option =0
Begin InputTables
    Name ="tblDesignCodeFurcationTube"
    Name ="SELECT Furcation.BaseID, Count(Furcation.BaseID) AS CountOfBaseID1 FROM (SELECT "
        "tblCableConstructions.BaseID, tblCableConstructions.AramidType1, tblCableConstru"
        "ctions.AramidEnds1, tblCableConstructions.AramidType2, tblCableConstructions.Ara"
        "midEnds2, tblCableConstructions.NominalWall, tblCableConstructions.NominalOD, tb"
        "lDesignCodeFurcationTube.TBNomOD, Count(tblFurcationTubeReference.baseid) AS Cou"
        "ntOfbaseid FROM tblDesignCodeFurcationTube INNER JOIN (tblFurcationTubeReference"
        " INNER JOIN tblCableConstructions ON tblFurcationTubeReference.BaseID = tblCable"
        "Constructions.BaseID) ON tblDesignCodeFurcationTube.Furcation = tblFurcationTube"
        "Reference.Furcation GROUP BY tblCableConstructions.BaseID, tblCableConstructions"
        ".AramidType1, tblCableConstructions.AramidEnds1, tblCableConstructions.AramidTyp"
        "e2, tblCableConstructions.AramidEnds2, tblCableConstructions.NominalWall, tblCab"
        "leConstructions.NominalOD, tblDesignCodeFurcationTube.TBNomOD)  AS Furcation GRO"
        "UP BY Furcation.BaseID HAVING (((Count(Furcation.BaseID))>1))"
    Alias ="MultipleFurcationSizes"
    Name ="tblFurcationTubeReference"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Expression ="tblCableConstructions.BaseID"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructions.NominalWall"
    Expression ="tblFurcationTubeReference.Furcation"
    Expression ="tblDesignCodeFurcationTube.TBNomOD"
    Expression ="tblDesignCodeFurcationTube.FurcationTube"
    Expression ="tblCableConstructions.AramidType1"
    Expression ="tblCableConstructions.AramidEnds1"
    Expression ="tblCableConstructions.AramidType2"
    Expression ="tblCableConstructions.AramidEnds2"
End
Begin Joins
    LeftTable ="MultipleFurcationSizes"
    RightTable ="tblFurcationTubeReference"
    Expression ="MultipleFurcationSizes.BaseID = tblFurcationTubeReference.BaseID"
    Flag =1
    LeftTable ="MultipleFurcationSizes"
    RightTable ="tblCableConstructions"
    Expression ="MultipleFurcationSizes.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblDesignCodeFurcationTube"
    RightTable ="tblFurcationTubeReference"
    Expression ="tblDesignCodeFurcationTube.Furcation = tblFurcationTubeReference.Furcation"
    Flag =1
End
Begin Groups
    Expression ="tblCableConstructions.BaseID"
    GroupLevel =0
    Expression ="tblCableConstructions.JacketMaterial"
    GroupLevel =0
    Expression ="tblCableConstructions.NominalOD"
    GroupLevel =0
    Expression ="tblCableConstructions.NominalWall"
    GroupLevel =0
    Expression ="tblFurcationTubeReference.Furcation"
    GroupLevel =0
    Expression ="tblDesignCodeFurcationTube.TBNomOD"
    GroupLevel =0
    Expression ="tblDesignCodeFurcationTube.FurcationTube"
    GroupLevel =0
    Expression ="tblCableConstructions.AramidType1"
    GroupLevel =0
    Expression ="tblCableConstructions.AramidEnds1"
    GroupLevel =0
    Expression ="tblCableConstructions.AramidType2"
    GroupLevel =0
    Expression ="tblCableConstructions.AramidEnds2"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x209dba2b6a73694c9f524e80c94e46d8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblFurcationTubeReference.Furcation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.FurcationTube"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.TBNomOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.BaseID"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =141
    Top =71
    Right =1603
    Bottom =457
    Left =-1
    Top =-1
    Right =1430
    Bottom =126
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =609
        Top =34
        Right =825
        Bottom =178
        Top =0
        Name ="tblDesignCodeFurcationTube"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MultipleFurcationSizes"
        Name =""
    End
    Begin
        Left =276
        Top =27
        Right =520
        Bottom =170
        Top =0
        Name ="tblFurcationTubeReference"
        Name =""
    End
    Begin
        Left =906
        Top =12
        Right =1311
        Bottom =178
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
