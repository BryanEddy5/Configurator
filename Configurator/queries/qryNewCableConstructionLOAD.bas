Operation =3
Name ="tblCableConstructions"
Option =0
Begin InputTables
    Name ="tblCableConstructions_Local"
End
Begin OutputColumns
    Name ="CablePasses"
    Expression ="tblCableConstructions_Local.CablePasses"
    Name ="StandardOperation"
    Expression ="tblCableConstructions_Local.StandardOperation"
    Name ="JacketMaterial"
    Expression ="tblCableConstructions_Local.JacketMaterial"
    Name ="NominalOD"
    Expression ="tblCableConstructions_Local.NominalOD"
    Name ="ODTolPlus"
    Expression ="tblCableConstructions_Local.ODTolPlus"
    Name ="ODTolMinus"
    Expression ="tblCableConstructions_Local.ODTolMinus"
    Name ="Height"
    Expression ="tblCableConstructions_Local.Height"
    Name ="HeightTolPlus"
    Expression ="tblCableConstructions_Local.HeightTolPlus"
    Name ="HeightTolMinus"
    Expression ="tblCableConstructions_Local.HeightTolMinus"
    Name ="NominalWall"
    Expression ="tblCableConstructions_Local.NominalWall"
    Name ="MaxAveWall"
    Expression ="tblCableConstructions_Local.MaxAveWall"
    Name ="MinAveWall"
    Expression ="tblCableConstructions_Local.MinAveWall"
    Name ="MinSpotWall"
    Expression ="tblCableConstructions_Local.MinSpotWall"
    Name ="AramidType1"
    Expression ="tblCableConstructions_Local.AramidType1"
    Name ="AramidEnds1"
    Expression ="tblCableConstructions_Local.AramidEnds1"
    Name ="AramidType2"
    Expression ="tblCableConstructions_Local.AramidType2"
    Name ="AramidEnds2"
    Expression ="tblCableConstructions_Local.AramidEnds2"
    Name ="ChipType"
    Expression ="tblCableConstructions_Local.ChipType"
    Name ="ColorChipPercentage"
    Expression ="tblCableConstructions_Local.ColorChipPercentage"
    Name ="Ripcord"
    Expression ="tblCableConstructions_Local.Ripcord"
    Name ="RipcordQty"
    Expression ="tblCableConstructions_Local.RipcordQty"
    Name ="LayLength"
    Expression ="tblCableConstructions_Local.LayLength"
    Name ="Talc"
    Expression ="tblCableConstructions_Local.Talc"
    Name ="FRP_Dia"
    Expression ="tblCableConstructions_Local.FRP_Dia"
    Name ="CM"
    Expression ="tblCableConstructions_Local.CM"
    Name ="CMMaterial"
    Expression ="tblCableConstructions_Local.CMMaterial"
    Name ="CM_OD"
    Expression ="tblCableConstructions_Local.CM_OD"
    Name ="CMODTolPlus"
    Expression ="tblCableConstructions_Local.CMODTolPlus"
    Name ="CMODTolMinus"
    Expression ="tblCableConstructions_Local.CMODTolMinus"
    Name ="CMWall"
    Expression ="tblCableConstructions_Local.CMWall"
    Name ="CMMinWall"
    Expression ="tblCableConstructions_Local.CMMinWall"
    Name ="HelixFactor"
    Expression ="tblCableConstructions_Local.HelixFactor"
    Name ="CoreWrap"
    Expression ="tblCableConstructions_Local.CoreWrap"
    Name ="CoreDia"
    Expression ="tblCableConstructions_Local.CoreDia"
    Name ="Binder1"
    Expression ="tblCableConstructions_Local.Binder1"
    Name ="QtyBinder1"
    Expression ="tblCableConstructions_Local.QtyBinder1"
    Name ="Binder1LayLength"
    Expression ="tblCableConstructions_Local.Binder1LayLength"
    Name ="Binder2"
    Expression ="tblCableConstructions_Local.Binder2"
    Name ="QtyBinder2"
    Expression ="tblCableConstructions_Local.QtyBinder2"
    Name ="Binder2LayLength"
    Expression ="tblCableConstructions_Local.Binder2LayLength"
    Name ="Binder3"
    Expression ="tblCableConstructions_Local.Binder3"
    Name ="QtyBinder3"
    Expression ="tblCableConstructions_Local.QtyBinder3"
    Name ="Binder3LayLength"
    Expression ="tblCableConstructions_Local.Binder3LayLength"
    Name ="ID"
    Expression ="tblCableConstructions_Local.ID"
    Name ="Additive"
    Expression ="tblCableConstructions_Local.Additive"
    Name ="AdditivePercentage"
    Expression ="tblCableConstructions_Local.AdditivePercentage"
    Name ="FirstRun"
    Expression ="tblCableConstructions_Local.FirstRun"
    Name ="EngineeringAssist"
    Expression ="tblCableConstructions_Local.EngineeringAssist"
    Name ="AramidType3"
    Expression ="tblCableConstructions_Local.AramidType3"
    Name ="AramidEnds3"
    Expression ="tblCableConstructions_Local.AramidEnds3"
    Name ="TemperatureApplication"
    Expression ="tblCableConstructions_Local.TemperatureApplication"
    Name ="TensileApplication"
    Expression ="tblCableConstructions_Local.TensileApplication"
    Name ="StandardsApplication"
    Expression ="tblCableConstructions_Local.StandardsApplication"
    Name ="Additive2"
    Expression ="tblCableConstructions_Local.Additive2"
    Name ="AdditivePercentage2"
    Expression ="tblCableConstructions_Local.AdditivePercentage2"
    Name ="TensileRatingShortTerm_N"
    Expression ="tblCableConstructions_Local.TensileRatingShortTerm_N"
    Name ="TensileRatingLongTerm_N"
    Expression ="tblCableConstructions_Local.TensileRatingLongTerm_N"
    Name ="EngineeringAssistReason"
    Expression ="tblCableConstructions_Local.EngineeringAssistReason"
    Name ="Instructions"
    Expression ="tblCableConstructions_Local.Instructions"
    Name ="Instructions2"
    Expression ="tblCableConstructions_Local.Instructions2"
    Name ="CableShapeID"
    Expression ="tblCableConstructions_Local.CableShapeID"
    Name ="FiberApplicationID"
    Expression ="tblCableConstructions_Local.FiberApplicationID"
    Name ="JacketStripID"
    Expression ="tblCableConstructions_Local.JacketStripID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x299524fce218ac4fbd9443ad073b51dd
End
Begin
    Begin
        dbText "Name" ="tblCableConstructions_Local.Instructions2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_Local.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_Local.CableShapeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_Local.FiberApplicationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_Local.JacketStripID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =8
    Right =1320
    Bottom =740
    Left =-1
    Top =-1
    Right =1260
    Bottom =520
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =299
        Top =48
        Right =755
        Bottom =388
        Top =0
        Name ="tblCableConstructions_Local"
        Name =""
    End
End
