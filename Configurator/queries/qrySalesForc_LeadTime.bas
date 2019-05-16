Operation =1
Option =0
Where ="(((tblCableConstructionReferences.Active)=True) AND ((tblCableConstructionRefere"
    "nces.CableFamily) Is Not Null) AND ((tblCableConstructionReferences.CableType) L"
    "ike \"finished*\"))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.LeadTime_ID"
    Expression ="tblCableConstructionReferences.ApprovedForQuoting"
    Expression ="tblCableConstructionReferences.FiberCount"
    Expression ="tblCableConstructions.CablePasses"
    Expression ="tblCableConstructionReferences.TBType"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="tblCableConstructionReferences.LeadTime_ID_Armored"
    Expression ="tblCableConstructionReferences.CableFamily"
    Expression ="tblCableConstructionReferences.CableLevel1"
    Expression ="tblCableConstructionReferences.CableType"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([qrySalesForc_LeadTime].[CableFamily]=\"u\")"
dbMemo "OrderBy" ="[qrySalesForc_LeadTime].[FiberCount], [qrySalesForc_LeadTime].[CableFamily], [qr"
    "ySalesForc_LeadTime].[CablePasses], [qrySalesForc_LeadTime].[LeadTime_ID], [qryS"
    "alesForc_LeadTime].[TBType], [qrySalesForc_LeadTime].[Base], [qrySalesForc_LeadT"
    "ime].[CableLevel1]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9120a492beb49244bd8beca827befe73
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ApprovedForQuoting"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FiberCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.LeadTime_ID_Armored"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =198
    Top =102
    Right =1307
    Bottom =731
    Left =-1
    Top =-1
    Right =1077
    Bottom =564
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =259
        Top =99
        Right =485
        Bottom =530
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =668
        Top =150
        Right =873
        Bottom =459
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
