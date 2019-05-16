dbMemo "SQL" ="SELECT tblCableConstructionReferences.DesignCode, tblCableConstructionReferences"
    ".Base, tblFurcationTubeReference.Furcation, tblFurcationTubeReference.BaseID, tb"
    "lDesignCodeFurcationTube.FurcationTube, tblCableConstructionReferences.ReleasedD"
    "esign, tblDesignCodeFurcationTube.TBNomOD\015\012FROM tblDesignCodeFurcationTube"
    " INNER JOIN ((tblCableConstructionReferences INNER JOIN tblFurcationTubeReferenc"
    "e ON tblCableConstructionReferences.BaseID = tblFurcationTubeReference.BaseID) I"
    "NNER JOIN dbo_tblCableConstructions ON tblCableConstructionReferences.BaseID = d"
    "bo_tblCableConstructions.BaseID) ON tblDesignCodeFurcationTube.Furcation = tblFu"
    "rcationTubeReference.Furcation\015\012WHERE (((tblCableConstructionReferences.Re"
    "leasedDesign)=True));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[qryFurcationTubeDescrapancey].[Furcation], [qryFurcationTubeDescrapancey].[Base"
    "], [qryFurcationTubeDescrapancey].[TBNomOD]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x08f2c7dbd09c2c47b8c5f4feb10cc855
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([qryFurcationTubeDescrapancey].[Base] Like \"D*\")"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblFurcationTubeReference.Furcation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFurcationTubeReference.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.FurcationTube"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFurcationTube.TBNomOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
    End
End
