Operation =1
Option =0
Begin InputTables
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Name ="dbo_CableUnion"
End
Begin OutputColumns
    Expression ="tblCableConstructions.StandardOperation"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructions.NominalOD"
    Alias ="UOM_Marking"
    Expression ="IIf([PrintLine4] Like \"*meter*\",\"M\",IIf([PrintLine4] Like \"*feet*\",\"FT\","
        "\"\"))"
    Alias ="RevDate"
    Expression ="Left([dbo_CableUnion].[Revision Date],10)"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructions.ODTolPlus"
    Expression ="dbo_CableUnion.Oracle"
    Expression ="dbo_CableUnion.Fiber"
    Expression ="dbo_CableUnion.Color"
    Expression ="dbo_CableUnion.Jacket"
    Expression ="dbo_CableUnion.Base"
    Expression ="dbo_CableUnion.ColorChip"
    Expression ="dbo_CableUnion.ListingCompany"
    Expression ="dbo_CableUnion.Customer"
    Expression ="dbo_CableUnion.Active"
    Expression ="dbo_CableUnion.Reason"
    Expression ="dbo_CableUnion.[Item No]"
    Expression ="dbo_CableUnion.Fiber2"
    Expression ="dbo_CableUnion.PID"
    Expression ="dbo_CableUnion.OracleStatus"
    Expression ="dbo_CableUnion.Enumber"
    Expression ="dbo_CableUnion.Fiber3"
    Expression ="dbo_CableUnion.CustomerRev"
    Expression ="dbo_CableUnion.Location"
    Expression ="dbo_CableUnion.RibbonHighCure"
    Expression ="dbo_CableUnion.OracleStatus"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="dbo_CableUnion"
    Expression ="tblCableConstructionReferences.Base = dbo_CableUnion.Base"
    Flag =1
End
Begin OrderBy
    Expression ="Left([dbo_CableUnion].[Revision Date],10)"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xc3dcb5cb36f68e449c7cfe99bd20cf0b
End
dbMemo "Filter" ="([qryfrmItemSearch].[Oracle] Like \"pt*\")"
Begin
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.ODTolPlus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UOM_Marking"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x47e8eead89df5c4c8bc0f0b88d7561ad
        End
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RevDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9d677e388df8a2429ec8294d21f11fe0
        End
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.PID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Enumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Fiber2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Fiber3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.RibbonHighCure"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1020"
        dbLong "AggregateType" ="-1"
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
    Bottom =269
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =821
        Top =19
        Right =1020
        Bottom =220
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =419
        Top =12
        Right =670
        Bottom =241
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =49
        Top =6
        Right =238
        Bottom =267
        Top =0
        Name ="dbo_CableUnion"
        Name =""
    End
End
