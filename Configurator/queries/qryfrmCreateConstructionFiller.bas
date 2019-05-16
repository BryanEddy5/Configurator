Operation =1
Option =0
Where ="(((tblCableConstructionReferences.CableType)=\"filler\") AND ((tblCableConstruct"
    "ionReferences.Active)=True))"
Begin InputTables
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.FiberCount"
    Expression ="tblCableConstructionReferences.FibersPerBundle"
    Expression ="tblCableConstructionReferences.TBType"
    Expression ="tblCableConstructionReferences.CableType"
    Expression ="tblCableConstructionReferences.Active"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbf17f1b5a204594ca93e04db165c0501
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FiberCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FibersPerBundle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =99
    Top =65
    Right =1139
    Bottom =809
    Left =-1
    Top =-1
    Right =1008
    Bottom =388
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =114
        Top =36
        Right =366
        Bottom =420
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
