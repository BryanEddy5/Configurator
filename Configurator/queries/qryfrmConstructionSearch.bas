Operation =1
Option =2
Begin InputTables
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Alias ="NomOD"
    Expression ="IIf([NominalOD] Is Null,0,CDbl([NominalOD]))"
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.DateCreated"
    Expression ="tblCableConstructionReferences.CreatedBy"
    Expression ="tblCableConstructionReferences.RevisionDate"
    Expression ="tblCableConstructionReferences.RevisedBy"
    Expression ="tblCableConstructionReferences.DesignCode"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Expression ="tblCableConstructionReferences.ConstructionDescription"
    Expression ="tblCableConstructionReferences.BaseID"
    Expression ="tblCableConstructionReferences.NumSubFillers"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="tblCableConstructionReferences.PrefixID"
    Expression ="tblCableConstructionReferences.ProductID"
    Expression ="tblCableConstructionReferences.SetupID"
    Expression ="tblCableConstructionReferences.FamilyID"
    Alias ="Expr1"
    Expression ="tblCableConstructionReferences.StandarOperation"
    Expression ="tblCableConstructionReferences.UnitIDTypeNumber"
    Expression ="tblCableConstructions.StandardOperation"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructions.ODTolPlus"
    Expression ="tblCableConstructions.ODTolMinus"
    Expression ="tblCableConstructions.Height"
    Expression ="tblCableConstructions.HeightTolPlus"
    Expression ="tblCableConstructions.HeightTolMinus"
    Expression ="tblCableConstructions.NominalWall"
    Expression ="tblCableConstructions.MaxAveWall"
    Expression ="tblCableConstructions.MinAveWall"
    Expression ="tblCableConstructions.MinSpotWall"
    Expression ="tblCableConstructions.AramidType1"
    Expression ="tblCableConstructions.AramidEnds1"
    Expression ="tblCableConstructions.AramidType2"
    Expression ="tblCableConstructions.AramidEnds2"
    Expression ="tblCableConstructions.ChipType"
    Expression ="tblCableConstructions.Ripcord"
    Expression ="tblCableConstructions.LayLength"
    Expression ="tblCableConstructions.Talc"
    Expression ="tblCableConstructions.FRP_Dia"
    Expression ="tblCableConstructions.CM"
    Expression ="tblCableConstructions.CMMaterial"
    Expression ="tblCableConstructions.CM_OD"
    Expression ="tblCableConstructions.CMODTolPlus"
    Expression ="tblCableConstructions.CMODTolMinus"
    Expression ="tblCableConstructions.CMWall"
    Expression ="tblCableConstructions.CMMinWall"
    Expression ="tblCableConstructions.HelixFactor"
    Expression ="tblCableConstructions.CoreWrap"
    Expression ="tblCableConstructions.CoreDia"
    Expression ="tblCableConstructions.Binder1"
    Expression ="tblCableConstructions.QtyBinder1"
    Expression ="tblCableConstructions.Binder1LayLength"
    Expression ="tblCableConstructions.Binder2"
    Expression ="tblCableConstructions.QtyBinder2"
    Expression ="tblCableConstructions.Binder2LayLength"
    Expression ="tblCableConstructions.Binder3"
    Expression ="tblCableConstructions.QtyBinder3"
    Expression ="tblCableConstructions.Binder3LayLength"
    Alias ="Location"
    Expression ="IIf([CablePasses] Is Null,0,CInt([CablePasses]))"
    Expression ="tblCableConstructions.RevisionDate"
    Expression ="tblCableConstructions.RevisionNumber"
    Expression ="tblCableConstructions.RevisionHistory"
    Expression ="tblCableConstructions.Instructions"
    Expression ="tblCableConstructions.Instructions2"
    Expression ="tblCableConstructions.SSMA_TimeStamp"
    Expression ="tblCableConstructions.DateCreated"
    Expression ="tblCableConstructions.CreatedBy"
    Expression ="tblCableConstructions.RevisedBy"
    Expression ="tblCableConstructions.ColorChipPercentage"
    Expression ="tblCableConstructionReferences.InactiveReason"
    Alias ="SubOD"
    Expression ="IIf([Location]=2,Mid([Base],7,3),\"\")"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructionReferences.NumSubPositions"
    Expression ="tblCableConstructionReferences.NumSubFillers"
    Expression ="tblCableConstructionReferences.ReleasedDesign"
    Expression ="tblCableConstructionReferences.FamilyID"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x327e7fcd666c374cba10fcbf2aa626fb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisionDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FamilyID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinSpotWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ColorChipPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.SetupID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.StandarOperation"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM_OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MaxAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CreatedBy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.PrefixID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DateCreated"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubFillers"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ProductID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.FRP_Dia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisedBy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ChipType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Ripcord"
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
        dbText "Name" ="tblCableConstructions.CMMinWall"
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
        dbText "Name" ="tblCableConstructions.CoreDia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1"
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
        dbText "Name" ="tblCableConstructions.RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionHistory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.SSMA_TimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x974976befc09fa49a1adca6e7634771c
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.InactiveReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="NomOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd4a6b21af8b91f45876d8d8cd1fe57d6
        End
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
    Bottom =212
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =948
        Top =14
        Right =1179
        Bottom =293
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =516
        Top =16
        Right =775
        Bottom =245
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
