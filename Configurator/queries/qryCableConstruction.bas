Operation =1
Option =0
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="vColor_Chip_Jacket_Material"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.*"
    Expression ="tblCableConstructions.StandardOperation"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructions.NominalOD"
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
    Expression ="tblCableConstructions.CablePasses"
    Expression ="tblCableConstructions.RevisionDate"
    Expression ="tblCableConstructions.RevisionNumber"
    Expression ="tblCableConstructions.RevisionHistory"
    Expression ="tblCableConstructions.Instructions"
    Expression ="tblCableConstructions.Instructions2"
    Expression ="tblCableConstructions.SSMA_TimeStamp"
    Expression ="tblCableConstructions.DateCreated"
    Expression ="tblCableConstructions.CreatedBy"
    Expression ="tblCableConstructions.BaseID"
    Expression ="tblCableConstructions.Additive"
    Expression ="tblCableConstructions.AdditivePercentage"
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryFiberOracle.Color"
    Expression ="qryFiberOracle.Jacket"
    Expression ="qryFiberOracle.Base"
    Expression ="qryFiberOracle.ListingCompany"
    Expression ="qryFiberOracle.Customer"
    Expression ="qryFiberOracle.PrintLine4"
    Expression ="qryFiberOracle.Active"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.ItemNo"
    Expression ="qryFiberOracle.Fiber2"
    Expression ="qryFiberOracle.PID"
    Expression ="qryFiberOracle.OracleStatus"
    Expression ="qryFiberOracle.Enumber"
    Expression ="qryFiberOracle.Fiber3"
    Expression ="qryFiberOracle.CustomerRev"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.[Revision Date]"
    Expression ="qryFiberOracle.RibbonHighCure"
    Expression ="qryFiberOracle.PrintSpacing"
    Expression ="qryFiberOracle.PrintLine2"
    Expression ="qryFiberOracle.DateCreated"
    Expression ="qryFiberOracle.PrintLine3"
    Expression ="qryFiberOracle.PrintLine1"
    Expression ="vColor_Chip_Jacket_Material.ColorChip"
    Expression ="vColor_Chip_Jacket_Material.Color_Chip_Concentration"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryFiberOracle.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
    LeftTable ="qryFiberOracle"
    RightTable ="vColor_Chip_Jacket_Material"
    Expression ="qryFiberOracle.Color = vColor_Chip_Jacket_Material.Color"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="vColor_Chip_Jacket_Material"
    Expression ="tblCableConstructions.JacketMaterial = vColor_Chip_Jacket_Material.JacketMateria"
        "l"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfc05c28133b6d04893dc4aabdde6496c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructions.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MaxAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.SetupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3450"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Base]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.PrefixID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Customer]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[OracleStatus]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Enumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ProductID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FamilyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
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
        dbText "Name" ="tblCableConstructions.CM_OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ChipType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMMaterial"
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
        dbText "Name" ="tblCableConstructions.LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM"
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
        dbText "Name" ="tblCableConstructions.Talc"
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
        dbText "Name" ="tblCableConstructions.Binder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2LayLength"
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
        dbText "Name" ="tblCableConstructions.CablePasses"
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
        dbText "Name" ="tblCableConstructionReferences.InactiveReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ColorChipPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Additive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AdditivePercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID_Armored"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FibersPerBundle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FiberCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumCopperUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Revision Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ApprovedForQuoting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.RibbonHighCure"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PrintLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vColor_Chip_Jacket_Material.ColorChip"
        dbInteger "ColumnWidth" ="4080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vColor_Chip_Jacket_Material.Color_Chip_Concentration"
        dbInteger "ColumnWidth" ="2865"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =37
    Top =73
    Right =1566
    Bottom =949
    Left =-1
    Top =-1
    Right =1497
    Bottom =418
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =35
        Right =322
        Bottom =345
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =442
        Top =34
        Right =755
        Bottom =378
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =823
        Top =22
        Right =1254
        Bottom =361
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =1303
        Top =89
        Right =1447
        Bottom =298
        Top =0
        Name ="vColor_Chip_Jacket_Material"
        Name =""
    End
End
