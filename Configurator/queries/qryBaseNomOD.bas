Operation =1
Option =0
Where ="(((tblCableConstructionReferences.Active)<>0))"
Begin InputTables
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Alias ="OD"
    Expression ="tblCableConstructions.NominalOD"
    Alias ="CabFam"
    Expression ="Left([tblCableConstructionReferences].[DesignCode],2)"
    Alias ="Jacket"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructionReferences.ConstructionDescription"
End
Begin Joins
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd0319cd6aea1df4fad21263633c4536f
End
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qryBaseNomOD].[Base]"
Begin
    Begin
        dbText "Name" ="CabFam"
        dbInteger "ColumnOrder" ="1"
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
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.[Base]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ConstructionDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb773cef5efc57b43b184667e5219a780
        End
    End
    Begin
        dbText "Name" ="ConstructionDescription"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf3e30d0aa95d24d87d6c74f18f82a3a
        End
    End
End
Begin
    State =0
    Left =126
    Top =118
    Right =1357
    Bottom =755
    Left =-1
    Top =-1
    Right =1199
    Bottom =198
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =139
        Top =21
        Right =316
        Bottom =195
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =398
        Top =33
        Right =689
        Bottom =177
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
