Operation =1
Option =0
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tblCableConstructions"
    Name ="qryBasePrint"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Alias ="UOM_Marking"
    Expression ="IIf(qryFiberOracle.PrintLine4 Like \"*meter*\",\"M\",IIf(qryFiberOracle.PrintLin"
        "e4 Like \"*feet*\",\"FT\",\"\"))"
    Alias ="RevDate"
    Expression ="Left(qryFiberOracle.[Revision Date],10)"
    Alias ="IsActive"
    Expression ="IIf(qryFiberOracle.Active=0,\"No\",\"Yes\")"
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryFiberOracle.Color"
    Expression ="qryFiberOracle.Jacket"
    Expression ="qryFiberOracle.ColorChip"
    Expression ="qryFiberOracle.ListingCompany"
    Expression ="qryFiberOracle.Customer"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Fiber2"
    Expression ="qryFiberOracle.PID"
    Expression ="qryFiberOracle.OracleStatus"
    Expression ="qryFiberOracle.Enumber"
    Expression ="qryFiberOracle.Fiber3"
    Expression ="qryFiberOracle.CustomerRev"
    Expression ="qryFiberOracle.Location"
    Expression ="qryFiberOracle.[Revision Date]"
    Expression ="tblCableConstructions.*"
    Expression ="qryBasePrint.PrintLine1"
    Expression ="qryBasePrint.PrintLine2"
    Expression ="qryBasePrint.PrintLine3"
    Expression ="qryBasePrint.PrintLine4"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="qryBasePrint"
    Expression ="qryFiberOracle.Oracle = qryBasePrint.Oracle"
    Flag =1
    LeftTable ="qryFiberOracle"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryFiberOracle.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
End
Begin OrderBy
    Expression ="Left(qryFiberOracle.[Revision Date],10)"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc65887c965bf9848ba09e95216eb76c4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="UOM_Marking"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb6cf3b588c9a454d9406f35c65a393c4
        End
    End
    Begin
        dbText "Name" ="RevDate"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8c1aa678f1ca1f4e94e6d5fd672e3c6f
        End
    End
    Begin
        dbText "Name" ="IsActive"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4261640f14ba0949a7a338e45f250aaf
        End
    End
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
        dbInteger "ColumnWidth" ="4635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Color"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Jacket"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ColorChip"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.ListingCompany"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="3270"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
        dbInteger "ColumnWidth" ="2100"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber2"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbInteger "ColumnOrder" ="4"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.OracleStatus"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Enumber"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber3"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.CustomerRev"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM"
        dbInteger "ColumnWidth" ="5715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionDate"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionNumber"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionHistory"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Revision Date]"
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
        dbText "Name" ="tblCableConstructions.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolMinus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ChipType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CoreDia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MaxAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinSpotWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1"
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
        dbText "Name" ="tblCableConstructions.Binder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Talc"
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
        dbText "Name" ="tblCableConstructions.FRP_Dia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMODTolPlus"
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
        dbText "Name" ="tblCableConstructions.Binder2LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.PrintLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.PrintLine4"
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
        dbText "Name" ="qryBasePrint.PrintLine2"
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
        dbText "Name" ="tblCableConstructions.Binder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.PrintLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions2"
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
        dbText "Name" ="tblCableConstructions.BaseID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1864
    Bottom =996
    Left =-1
    Top =-1
    Right =1832
    Bottom =162
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =739
        Top =22
        Right =1034
        Bottom =200
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =247
        Top =18
        Right =453
        Bottom =196
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =1142
        Top =10
        Right =1341
        Bottom =173
        Top =0
        Name ="qryBasePrint"
        Name =""
    End
    Begin
        Left =553
        Top =37
        Right =697
        Bottom =181
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
