Operation =1
Option =4
Where ="(((tblCableConstructionReferences.TBType)=\"furcation\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.TBType"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryFiberOracle.Base = tblCableConstructionReferences.Base"
    Flag =1
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
    0xf1003fb57bea384ea71156683c80c481
End
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =178
    Right =1587
    Bottom =967
    Left =-1
    Top =-1
    Right =1555
    Bottom =435
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =261
        Bottom =220
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =399
        Top =19
        Right =596
        Bottom =222
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End
