Operation =1
Option =0
Where ="(((tblCableConstructionReferences.Active)<>0) AND (([Product Desrcriptions].SubF"
    "iller)=False) AND ((tblDesignTypeGroups.Active)<>0) AND ((tblDesignTypes.Active)"
    "<>0))"
Begin InputTables
    Name ="tblDesignTypeGroups"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Name ="Product Desrcriptions"
    Name ="tblDesignTypes"
    Name ="tblCableTightBufferReference"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Alias ="SubOD"
    Expression ="Mid([DesignCode],6,1) & \".\" & Mid([DesignCode],7,1)"
    Alias ="SubFiber"
    Expression ="Mid([Level 1 Product],3,3)"
    Alias ="NumPositions"
    Expression ="tblCableConstructionReferences.NumSubPositions"
    Alias ="FiberCount"
    Expression ="Mid([DesignCode],3,3)"
    Expression ="tblCableConstructionReferences.ConstructionDescription"
    Alias ="Jacket"
    Expression ="tblCableConstructions.JacketMaterial"
    Alias ="OD"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructionReferences.BaseID"
    Expression ="tblCableConstructionReferences.DesignCode"
    Alias ="CabFam"
    Expression ="Left([tblCableConstructionReferences].[DesignCode],2)"
    Expression ="tblCableConstructions.CablePasses"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="[Product Desrcriptions].SubFiller"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Expression ="tblCableConstructionReferences.CableFamily"
    Expression ="tblDesignTypes.DesignTypeID"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID"
    Expression ="tblDesignTypeGroups.Active"
    Expression ="tblDesignTypes.Active"
    Expression ="tblDesignTypeGroups.DesignTypeGroupDesc"
    Expression ="tblDesignTypes.DesignTypeDesc"
    Expression ="tblCableConstructionReferences.CableLevel1"
    Expression ="tblCableTightBufferReference.TBType"
End
Begin Joins
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignTypes"
    Expression ="tblCableConstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID"
    Flag =1
    LeftTable ="tblDesignTypeGroups"
    RightTable ="tblDesignTypes"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID = tblDesignTypes.DesignTypeGroupID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableConstructionReferences.TBType = tblCableTightBufferReference.TBType"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
    Flag =1
End
Begin OrderBy
    Expression ="Mid([DesignCode],6,1) & \".\" & Mid([DesignCode],7,1)"
    Flag =0
    Expression ="Mid([Level 1 Product],3,3)"
    Flag =0
    Expression ="tblCableConstructionReferences.NumSubPositions"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xefb9f104c2a68e49ada0b8797c1506fe
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CabFam"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdaa6e18a2256364686bf5d145ab90c28
        End
    End
    Begin
        dbText "Name" ="NumPositions"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1195c0264c858c45bb61e7e0650e3316
        End
    End
    Begin
        dbText "Name" ="OD"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfd8dce40c58ada46a931db3d248c8369
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.[Base]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jacket"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2360a1788b363e4392e02bded125fbd8
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FiberCount"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb140eb905090fc4e8e6ed6de7e315d9d
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.[Active]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubFiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbInteger "ColumnWidth" ="3780"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubOD"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1e54e59c29c94f498a25064a28091285
        End
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].SubFiller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.Active"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblDesignTypeGroups.DesignTypeGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypeGroups.Active"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3060"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblDesignTypes.DesignTypeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypeGroups.DesignTypeGroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.TBType"
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
    Bottom =443
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1327
        Top =85
        Right =1530
        Bottom =244
        Top =0
        Name ="tblDesignTypeGroups"
        Name =""
    End
    Begin
        Left =103
        Top =12
        Right =309
        Bottom =280
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =350
        Top =4
        Right =649
        Bottom =428
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =753
        Top =25
        Right =1013
        Bottom =448
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =1046
        Top =9
        Right =1231
        Bottom =231
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
    Begin
        Left =1514
        Top =14
        Right =1726
        Bottom =196
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End
