Operation =3
Name ="tblCableConstructions_Local"
Option =0
Where ="(((tblCableConstructionReferences.Base)=[Forms]![frmCreateConstruction]![cbBase_"
    "Copy]))"
Begin InputTables
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Name ="CablePasses"
    Expression ="tblCableConstructions.CablePasses"
    Name ="StandardOperation"
    Expression ="tblCableConstructions.StandardOperation"
    Name ="JacketMaterial"
    Expression ="tblCableConstructions.JacketMaterial"
    Name ="NominalOD"
    Expression ="tblCableConstructions.NominalOD"
    Name ="ODTolPlus"
    Expression ="tblCableConstructions.ODTolPlus"
    Name ="ODTolMinus"
    Expression ="tblCableConstructions.ODTolMinus"
    Name ="Height"
    Expression ="tblCableConstructions.Height"
    Name ="HeightTolPlus"
    Expression ="tblCableConstructions.HeightTolPlus"
    Name ="HeightTolMinus"
    Expression ="tblCableConstructions.HeightTolMinus"
    Name ="NominalWall"
    Expression ="tblCableConstructions.NominalWall"
    Name ="MaxAveWall"
    Expression ="tblCableConstructions.MaxAveWall"
    Name ="MinAveWall"
    Expression ="tblCableConstructions.MinAveWall"
    Name ="MinSpotWall"
    Expression ="tblCableConstructions.MinSpotWall"
    Name ="AramidType1"
    Expression ="tblCableConstructions.AramidType1"
    Name ="AramidEnds1"
    Expression ="tblCableConstructions.AramidEnds1"
    Name ="AramidType2"
    Expression ="tblCableConstructions.AramidType2"
    Name ="AramidEnds2"
    Expression ="tblCableConstructions.AramidEnds2"
    Name ="ChipType"
    Expression ="tblCableConstructions.ChipType"
    Name ="ColorChipPercentage"
    Expression ="tblCableConstructions.ColorChipPercentage"
    Name ="Ripcord"
    Expression ="tblCableConstructions.Ripcord"
    Name ="RipcordQty"
    Expression ="tblCableConstructions.RipcordQty"
    Name ="LayLength"
    Expression ="tblCableConstructions.LayLength"
    Name ="Talc"
    Expression ="tblCableConstructions.Talc"
    Name ="FRP_Dia"
    Expression ="tblCableConstructions.FRP_Dia"
    Name ="CM"
    Expression ="tblCableConstructions.CM"
    Name ="CMMaterial"
    Expression ="tblCableConstructions.CMMaterial"
    Name ="CM_OD"
    Expression ="tblCableConstructions.CM_OD"
    Name ="CMODTolPlus"
    Expression ="tblCableConstructions.CMODTolPlus"
    Name ="CMODTolMinus"
    Expression ="tblCableConstructions.CMODTolMinus"
    Name ="CMWall"
    Expression ="tblCableConstructions.CMWall"
    Name ="CMMinWall"
    Expression ="tblCableConstructions.CMMinWall"
    Name ="HelixFactor"
    Expression ="tblCableConstructions.HelixFactor"
    Name ="CoreWrap"
    Expression ="tblCableConstructions.CoreWrap"
    Name ="CoreDia"
    Expression ="tblCableConstructions.CoreDia"
    Name ="Binder1"
    Expression ="tblCableConstructions.Binder1"
    Name ="QtyBinder1"
    Expression ="tblCableConstructions.QtyBinder1"
    Name ="Binder1LayLength"
    Expression ="tblCableConstructions.Binder1LayLength"
    Name ="Binder2"
    Expression ="tblCableConstructions.Binder2"
    Name ="QtyBinder2"
    Expression ="tblCableConstructions.QtyBinder2"
    Name ="Binder2LayLength"
    Expression ="tblCableConstructions.Binder2LayLength"
    Name ="Binder3"
    Expression ="tblCableConstructions.Binder3"
    Name ="QtyBinder3"
    Expression ="tblCableConstructions.QtyBinder3"
    Name ="Binder3LayLength"
    Expression ="tblCableConstructions.Binder3LayLength"
    Name ="Additive"
    Expression ="tblCableConstructions.Additive"
    Name ="AdditivePercentage"
    Expression ="tblCableConstructions.AdditivePercentage"
    Name ="EngineeringAssist"
    Expression ="tblCableConstructions.EngineeringAssist"
    Name ="AramidType3"
    Expression ="tblCableConstructions.AramidType3"
    Name ="AramidEnds3"
    Expression ="tblCableConstructions.AramidEnds3"
    Name ="TemperatureApplication"
    Expression ="tblCableConstructions.TemperatureApplication"
    Name ="StandardsApplication"
    Expression ="tblCableConstructions.StandardsApplication"
    Name ="Additive2"
    Expression ="tblCableConstructions.Additive2"
    Name ="AdditivePercentage2"
    Expression ="tblCableConstructions.AdditivePercentage2"
    Name ="TensileRatingShortTerm_N"
    Expression ="tblCableConstructions.TensileRatingShortTerm_N"
    Name ="TensileRatingLongTerm_N"
    Expression ="tblCableConstructions.TensileRatingLongTerm_N"
    Name ="EngineeringAssistReason"
    Expression ="tblCableConstructions.EngineeringAssistReason"
    Name ="Instructions"
    Expression ="tblCableConstructions.Instructions"
    Name ="Instructions2"
    Expression ="tblCableConstructions.Instructions2"
    Name ="CableShapeID"
    Expression ="tblCableConstructions.CableShapeID"
    Name ="FiberApplicationID"
    Expression ="tblCableConstructions.FiberApplicationID"
    Name ="JacketStripID"
    Expression ="tblCableConstructions.JacketStripID"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x76e5d2fa0187654d883df98b50f7b465
End
Begin
    Begin
        dbText "Name" ="tblCableConstructions.StandardsApplication"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.TensileApplication"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Additive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionHistory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.EngineeringAssistReason"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MaxAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ChipType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RipcordQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ColorChipPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM_OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.FRP_Dia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinSpotWall"
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
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMODTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMODTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AdditivePercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Additive2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.TensileRatingLongTerm_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMMinWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AdditivePercentage2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HelixFactor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CoreWrap"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.FirstRun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.TemperatureApplication"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CoreDia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.EngineeringAssist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.TensileRatingShortTerm_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CableShapeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.FiberApplicationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketStripID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =124
    Top =56
    Right =1524
    Bottom =861
    Left =-1
    Top =-1
    Right =1368
    Bottom =368
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =565
        Top =7
        Right =1042
        Bottom =387
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =207
        Top =17
        Right =376
        Bottom =366
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
