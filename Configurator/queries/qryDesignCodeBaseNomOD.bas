Operation =1
Option =2
Where ="(((tblCableConstructionReferences.Base) Not Like \"ia*\") And ((tblCableConstruc"
    "tionReferences.FiberCount) Like Forms!frmDesignCode!FiberCount & \"*\") And ((tb"
    "lCableConstructionReferences.Active)<>0) And (([Product Desrcriptions].SubFiller"
    ")=False) And ((tblDesignTypeGroups.DesignTypeGroupID) Like Forms!frmDesignCode!C"
    "ableType & \"*\") And ((tblDesignTypes.SalesForceDescription) Like Forms!frmDesi"
    "gnCode!Jacket & \"*\") And ((tblDesignCodeJacket.JacketDescription_SalesForce) L"
    "ike fORMS!frmDesignCode!CableRating & \"*\"))"
Begin InputTables
    Name ="tblDesignTypeGroups"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Name ="tblDesignTypes"
    Name ="tblDesignCodeJacket"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.FiberCount"
    Alias ="SubOD"
    Expression ="Mid([DesignCode],6,1) & \".\" & Mid([DesignCode],7,1)"
    Alias ="SubFiber"
    Expression ="Mid([Level 1 Product],3,3)"
    Alias ="NumPositions"
    Expression ="tblCableConstructionReferences.NumSubPositions"
    Expression ="tblCableConstructionReferences.ConstructionDescription"
    Alias ="Jacket"
    Expression ="tblCableConstructions.JacketMaterial"
    Alias ="OD"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Expression ="tblCableConstructionReferences.BaseID"
    Expression ="tblCableConstructionReferences.DesignCode"
    Alias ="CabFam"
    Expression ="Left([tblCableConstructionReferences].[DesignCode],2)"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="[Product Desrcriptions].SubFiller"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID"
    Expression ="tblDesignTypes.DesignTypeID"
    Expression ="tblDesignCodeJacket.JacketLetter"
    Expression ="tblDesignTypes.SalesForceDescription"
    Expression ="tblDesignCodeJacket.JacketDescription_SalesForce"
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
    RightTable ="tblDesignCodeJacket"
    Expression ="tblCableConstructionReferences.CableLevel1 = tblDesignCodeJacket.JacketLetter"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignCodeJacket"
    Expression ="tblCableConstructionReferences.CableFamily = tblDesignCodeJacket.CableLetter"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
    Flag =2
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
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xefb9f104c2a68e49ada0b8797c1506fe
End
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
        dbText "Name" ="OD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfd8dce40c58ada46a931db3d248c8369
        End
        dbInteger "ColumnOrder" ="1"
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.[Base]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.[Active]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
        dbInteger "ColumnWidth" ="3780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubOD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
        dbBinary "GUID" = Begin
            0x6c8bcfac3f56324c9ba74125ca09bf0f
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.BaseID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SubFiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="NumPositions"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06004ebd83dfd748949fb235763ec340
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
        dbText "Name" ="tblDesignTypeGroups.DesignTypeGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacket.JacketLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FiberCount"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.SalesForceDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacket.JacketDescription_SalesForce"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =85
    Top =68
    Right =1788
    Bottom =811
    Left =-1
    Top =-1
    Right =1671
    Bottom =259
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1450
        Top =22
        Right =1631
        Bottom =166
        Top =0
        Name ="tblDesignTypeGroups"
        Name =""
    End
    Begin
        Left =103
        Top =12
        Right =317
        Bottom =186
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =380
        Top =11
        Right =717
        Bottom =308
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1240
        Top =14
        Right =1431
        Bottom =158
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
    Begin
        Left =1068
        Top =10
        Right =1292
        Bottom =156
        Top =0
        Name ="tblDesignCodeJacket"
        Name =""
    End
    Begin
        Left =760
        Top =12
        Right =1020
        Bottom =170
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End
