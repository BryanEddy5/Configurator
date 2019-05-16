Operation =1
Option =0
Where ="(((dbo_vCreateItemBaseCenterMember.DesignCode) Like Left([Forms]![frmMainMenu]!["
    "ItemNo],5) & Mid([Forms]![frmMainMenu]![ItemNo],7,2)))"
Begin InputTables
    Name ="dbo_vCreateItemBaseCenterMember"
End
Begin OutputColumns
    Expression ="dbo_vCreateItemBaseCenterMember.Base"
    Expression ="dbo_vCreateItemBaseCenterMember.ConstructionDescription"
    Expression ="dbo_vCreateItemBaseCenterMember.OD"
    Expression ="dbo_vCreateItemBaseCenterMember.NumPositions"
    Expression ="dbo_vCreateItemBaseCenterMember.NumFillers"
    Expression ="dbo_vCreateItemBaseCenterMember.JacketMaterial"
    Expression ="dbo_vCreateItemBaseCenterMember.CountOfBase"
    Expression ="dbo_vCreateItemBaseCenterMember.Active"
    Expression ="dbo_vCreateItemBaseCenterMember.DesignCode"
    Expression ="dbo_vCreateItemBaseCenterMember.Location"
    Expression ="dbo_vCreateItemBaseCenterMember.ReleasedDesign"
    Expression ="dbo_vCreateItemBaseCenterMember.DesignExtension"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x61bf44b0a8c5cb4c9b08fc728f27f852
End
dbMemo "Filter" ="([qryfrmCreateItemBaseCenterMember].[Base] Like \"M*\")"
Begin
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.CountOfBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.NumPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.NumFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.ConstructionDescription"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vCreateItemBaseCenterMember.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DesignExtension"
        dbBinary "GUID" = Begin
            0xfd6d1f65a5da3246bca2d1a774c2b137
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1551
    Bottom =957
    Left =-1
    Top =-1
    Right =1519
    Bottom =207
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =251
        Top =9
        Right =484
        Bottom =153
        Top =0
        Name ="dbo_vCreateItemBaseCenterMember"
        Name =""
    End
End
