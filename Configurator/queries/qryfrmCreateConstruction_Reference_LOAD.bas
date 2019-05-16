Operation =3
Name ="tblCableConstructionReferences_Local"
Option =0
Where ="(((tblCableConstructionReferences.Base)=[Forms]![frmCreateConstruction]![cbBase_"
    "Copy]))"
Begin InputTables
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Name ="DesignCode"
    Expression ="tblCableConstructionReferences.DesignCode"
    Name ="DesignExtension"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Name ="ConstructionDescription"
    Expression ="tblCableConstructionReferences.ConstructionDescription"
    Name ="BaseID"
    Expression ="tblCableConstructionReferences.BaseID"
    Name ="NumSubFillers"
    Expression ="tblCableConstructionReferences.NumSubFillers"
    Name ="PrefixID"
    Expression ="tblCableConstructionReferences.PrefixID"
    Name ="FamilyID"
    Expression ="tblCableConstructionReferences.FamilyID"
    Name ="FibersPerBundle"
    Expression ="tblCableConstructionReferences.FibersPerBundle"
    Name ="UnitIDTypeNumber"
    Expression ="tblCableConstructionReferences.UnitIDTypeNumber"
    Name ="NumSubPositions"
    Expression ="tblCableConstructionReferences.NumSubPositions"
    Name ="TBType"
    Expression ="tblCableConstructionReferences.TBType"
    Name ="CableFamily"
    Expression ="tblCableConstructionReferences.CableFamily"
    Name ="CableLevel1"
    Expression ="tblCableConstructionReferences.CableLevel1"
    Name ="DesignTypeID"
    Expression ="tblCableConstructionReferences.DesignTypeID"
    Name ="NumCopperUnits"
    Expression ="tblCableConstructionReferences.NumCopperUnits"
    Name ="FiberCount"
    Expression ="tblCableConstructionReferences.FiberCount"
    Name ="CableType"
    Expression ="tblCableConstructionReferences.CableType"
    Name ="LeadTime_ID"
    Expression ="tblCableConstructionReferences.LeadTime_ID"
    Name ="LeadTime_ID_Armored"
    Expression ="tblCableConstructionReferences.LeadTime_ID_Armored"
    Name ="ImageGroupID"
    Expression ="tblCableConstructionReferences.ImageGroupID"
    Name ="Base"
    Expression ="tblCableConstructionReferences.Base"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc632b71903e1394db230e2de99dce05c
End
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID_Armored"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FamilyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.InactiveReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.PrefixID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ApprovedForQuoting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ProductID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.SetupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FiberCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FibersPerBundle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumCopperUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ImageGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1025"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.IsCommoned"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-323
    Top =58
    Right =1237
    Bottom =872
    Left =-1
    Top =-1
    Right =1528
    Bottom =355
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =275
        Top =116
        Right =581
        Bottom =351
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
