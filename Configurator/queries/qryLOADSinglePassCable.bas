Operation =3
Name ="tblSinglePassCable"
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #])=[Forms]![frmfinditem]![Orac"
    "le]))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Name ="New Oracle Part #"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Name ="Item No"
    Expression ="[Basic Product Construction].[Item No]"
    Name ="Customer"
    Expression ="[Basic Product Construction].Customer"
    Name ="Customer Part#"
    Expression ="[Basic Product Construction].[Customer Part#]"
    Name ="CustomerRev"
    Expression ="[Basic Product Construction].CustomerRev"
    Name ="Reason"
    Expression ="[Basic Product Construction].Reason"
    Name ="calculate bill?"
    Expression ="[Basic Product Construction].[calculate bill?]"
    Name ="Base"
    Expression ="[Basic Product Construction].Base"
    Name ="Cable Type"
    Expression ="[Basic Product Construction].[Cable Type]"
    Name ="Revision Date"
    Expression ="[Basic Product Construction].[Revision Date]"
    Name ="Revision Letter"
    Expression ="[Basic Product Construction].[Revision Letter]"
    Name ="Revision History"
    Expression ="[Basic Product Construction].[Revision History]"
    Name ="Standard Operation"
    Expression ="[Basic Product Construction].[Standard Operation]"
    Name ="Jacket Material"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Name ="Jacket Color"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Name ="Color Chip ID"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Name ="Print Reel No"
    Expression ="[Basic Product Construction].[Print Reel No]"
    Name ="Print Item No"
    Expression ="[Basic Product Construction].[Print Item No]"
    Name ="Std Subunit Print"
    Expression ="[Basic Product Construction].[Std Subunit Print]"
    Name ="Print Type (base)"
    Expression ="[Basic Product Construction].[Print Type (base)]"
    Name ="Print Line 1"
    Expression ="[Basic Product Construction].[Print Line 1]"
    Name ="Print Line 2"
    Expression ="[Basic Product Construction].[Print Line 2]"
    Name ="Print Line 3"
    Expression ="[Basic Product Construction].[Print Line 3]"
    Name ="Print Line 4"
    Expression ="[Basic Product Construction].[Print Line 4]"
    Name ="NewPrintLine4"
    Expression ="[Basic Product Construction].NewPrintLine4"
    Name ="Print Spacing"
    Expression ="[Basic Product Construction].[Print Spacing]"
    Name ="PrintNotes"
    Expression ="[Basic Product Construction].PrintNotes"
    Name ="Listing Company"
    Expression ="[Basic Product Construction].[Listing Company]"
    Name ="UL-ETL-Listing"
    Expression ="[Basic Product Construction].[UL-ETL-Listing]"
    Name ="UL-ETL-Const"
    Expression ="[Basic Product Construction].[UL-ETL-Const]"
    Name ="UL-ETL-Section"
    Expression ="[Basic Product Construction].[UL-ETL-Section]"
    Name ="ENumber"
    Expression ="[Basic Product Construction].ENumber"
    Name ="Label Type"
    Expression ="[Basic Product Construction].[Label Type]"
    Name ="EZ Strip"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Name ="TB Material"
    Expression ="[Basic Product Construction].[TB Material]"
    Name ="TB Nominal OD"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Name ="TB OD Tol  (+)"
    Expression ="[Basic Product Construction].[TB OD Tol  (+)]"
    Name ="TB OD Tol  (-)"
    Expression ="[Basic Product Construction].[TB OD Tol  (-)]"
    Name ="TB Color Series"
    Expression ="[Basic Product Construction].[TB Color Series]"
    Name ="Fiber Type"
    Expression ="[Basic Product Construction].[Fiber Type]"
    Name ="Fiber Oracle item"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Name ="FiberType2"
    Expression ="[Basic Product Construction].FiberType2"
    Name ="FiberType3"
    Expression ="[Basic Product Construction].FiberType3"
    Name ="1st Req Freq"
    Expression ="[Basic Product Construction].[1st Req Freq]"
    Name ="1st Max Atten"
    Expression ="[Basic Product Construction].[1st Max Atten]"
    Name ="1st Min BandW"
    Expression ="[Basic Product Construction].[1st Min BandW]"
    Name ="2nd Req Freq"
    Expression ="[Basic Product Construction].[2nd Req Freq]"
    Name ="2nd Max Atten"
    Expression ="[Basic Product Construction].[2nd Max Atten]"
    Name ="2nd Min BandW"
    Expression ="[Basic Product Construction].[2nd Min BandW]"
    Name ="SM 1300 Max Atten"
    Expression ="[Basic Product Construction].[SM 1300 Max Atten]"
    Name ="SM 1550 Max Atten"
    Expression ="[Basic Product Construction].[SM 1550 Max Atten]"
    Name ="Special Instr Product1"
    Expression ="[Basic Product Construction].[Special Instr Product1]"
    Name ="zzSpecial Instr Product3"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product3]"
    Name ="zzSpecial Instr Product4"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product4]"
    Name ="zzOracle Part#"
    Expression ="[Basic Product Construction].[zzOracle Part#]"
    Name ="OracleStatus"
    Expression ="[Basic Product Construction].OracleStatus"
    Name ="DateCreated"
    Expression ="[Basic Product Construction].DateCreated"
    Name ="CreatedBy"
    Expression ="[Basic Product Construction].CreatedBy"
    Name ="RevisedBy"
    Expression ="[Basic Product Construction].RevisedBy"
    Name ="RevisedDate"
    Expression ="[Basic Product Construction].RevisedDate"
    Name ="DesignCode"
    Expression ="tblCableConstructionReferences.DesignCode"
    Name ="DesignExtension"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Name ="BaseID"
    Expression ="tblCableConstructionReferences.BaseID"
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
    Name ="Ripcord"
    Expression ="tblCableConstructions.Ripcord"
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
    Name ="CablePasses"
    Expression ="tblCableConstructions.CablePasses"
    Name ="RevisionDate"
    Expression ="tblCableConstructions.RevisionDate"
    Name ="RevisionNumber"
    Expression ="tblCableConstructions.RevisionNumber"
    Name ="RevisionHistory"
    Expression ="tblCableConstructions.RevisionHistory"
    Name ="Instructions"
    Expression ="tblCableConstructions.Instructions"
    Name ="Instructions2"
    Expression ="tblCableConstructions.Instructions2"
    Alias ="Expr1"
    Name ="NumSubs"
    Expression ="tblCableConstructions.NumSubs"
    Alias ="Location"
    Name ="Location"
    Expression ="1"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =2
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x2df54c83fc49514aa29e9f56f4a0e934
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].RevisedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Height Tol (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NumSubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Cable Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Standard Operation]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[calculate bill?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision History]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Std Subunit Print]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Type (base)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[SM 1300 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzSpecial Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[SM 1550 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.MinAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzOracle Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.DesignCode"
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
        dbText "Name" ="tblCableConstructions.NominalWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ChipType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CoreDia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
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
        dbText "Name" ="tblCableConstructions.MaxAveWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.QtyBinder2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.HeightTolMinus"
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
        dbText "Name" ="tblCableConstructions.MinSpotWall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder1"
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
        dbText "Name" ="tblCableConstructions.QtyBinder3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CMMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.FRP_Dia"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CM_OD"
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
        dbText "Name" ="tblCableConstructions.Binder1LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Binder3LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
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
        dbText "Name" ="tblCableConstructions.RevisionNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4513b386eaf2c44da5b2cdf3e1562a51
        End
    End
    Begin
        dbText "Name" ="tblCableConstructions.RevisionHistory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.Instructions2"
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
            0x00000000000000000000000001000000
        End
    End
End
Begin
    State =0
    Left =13
    Top =43
    Right =1305
    Bottom =476
    Left =-1
    Top =-1
    Right =1260
    Bottom =131
    Left =1526
    Top =0
    ColumnsShown =651
    Begin
        Left =-1478
        Top =12
        Right =-1290
        Bottom =156
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =-1138
        Top =28
        Right =-994
        Bottom =172
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =-834
        Top =19
        Right =-609
        Bottom =163
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
