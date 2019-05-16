Operation =3
Name ="tblCableConstructionReferences"
Option =0
Begin InputTables
    Name ="tblCableConstructionReferences_Local"
End
Begin OutputColumns
    Name ="Base"
    Expression ="tblCableConstructionReferences_Local.Base"
    Name ="DesignCode"
    Expression ="tblCableConstructionReferences_Local.DesignCode"
    Name ="DesignExtension"
    Expression ="tblCableConstructionReferences_Local.DesignExtension"
    Name ="ConstructionDescription"
    Expression ="tblCableConstructionReferences_Local.ConstructionDescription"
    Name ="BaseID"
    Expression ="tblCableConstructionReferences_Local.BaseID"
    Name ="NumSubFillers"
    Expression ="tblCableConstructionReferences_Local.NumSubFillers"
    Name ="PrefixID"
    Expression ="tblCableConstructionReferences_Local.PrefixID"
    Name ="FamilyID"
    Expression ="tblCableConstructionReferences_Local.FamilyID"
    Name ="FibersPerBundle"
    Expression ="tblCableConstructionReferences_Local.FibersPerBundle"
    Name ="UnitIDTypeNumber"
    Expression ="tblCableConstructionReferences_Local.UnitIDTypeNumber"
    Name ="NumSubPositions"
    Expression ="tblCableConstructionReferences_Local.NumSubPositions"
    Name ="TBType"
    Expression ="tblCableConstructionReferences_Local.TBType"
    Name ="CableFamily"
    Expression ="tblCableConstructionReferences_Local.CableFamily"
    Name ="CableLevel1"
    Expression ="tblCableConstructionReferences_Local.CableLevel1"
    Name ="DesignTypeID"
    Expression ="tblCableConstructionReferences_Local.DesignTypeID"
    Name ="NumCopperUnits"
    Expression ="tblCableConstructionReferences_Local.NumCopperUnits"
    Name ="FiberCount"
    Expression ="tblCableConstructionReferences_Local.FiberCount"
    Name ="CableType"
    Expression ="tblCableConstructionReferences_Local.CableType"
    Name ="LeadTime_ID"
    Expression ="tblCableConstructionReferences_Local.LeadTime_ID"
    Name ="LeadTime_ID_Armored"
    Expression ="tblCableConstructionReferences_Local.LeadTime_ID_Armored"
    Name ="ImageGroupID"
    Expression ="tblCableConstructionReferences_Local.ImageGroupID"
    Alias ="Expr2"
    Name ="ReleasedDesign"
    Expression ="[Forms]![frmCreateConstruction]![cbTest]"
    Alias ="Expr1"
    Name ="IsCommoned"
    Expression ="[Forms]![frmCreateConstruction]![cbIsCommoned]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x8318eed9456ef5429cf5e38f8b946e70
End
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.InactiveReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.NumSubPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.PrefixID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.IsCommoned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9205760a8d70e1478addd81a4a01daa9
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.NumSubFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.FamilyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.ImageGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.FibersPerBundle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.FiberCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.CableType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x538004b3649aec46bf35ee67ca644a46
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.ConstructionDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.NumCopperUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.LeadTime_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_Local.LeadTime_ID_Armored"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =149
    Right =1333
    Bottom =853
    Left =-1
    Top =-1
    Right =1260
    Bottom =398
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =231
        Top =84
        Right =1096
        Bottom =502
        Top =0
        Name ="tblCableConstructionReferences_Local"
        Name =""
    End
End
