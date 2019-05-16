Operation =1
Option =0
Where ="((([Basic Product Construction].[Item No]) Not Like \"T?001*\"))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Expression ="[Basic Product Construction].[TB Chips Type]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd9acf673dba54143b0bc5c1633e3b89f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.New Oracle Part #"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.EZ Strip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Chips Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =459
    Top =99
    Right =1952
    Bottom =885
    Left =-1
    Top =-1
    Right =1461
    Bottom =418
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =510
        Bottom =268
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =379
        Top =302
        Right =698
        Bottom =472
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =822
        Top =114
        Right =1115
        Bottom =327
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
