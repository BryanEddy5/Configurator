Operation =1
Option =0
Where ="(((tblCableConstructions.StandardOperation) Is Null))"
Begin InputTables
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Expression ="tblCableConstructions.StandardOperation"
    Expression ="tblCableConstructions.BaseID"
End
Begin OrderBy
    Expression ="tblCableConstructions.StandardOperation"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0a52233347010348bb8c28fc40030789
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[StandardOperation]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[BaseID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.[StandardOperation]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.[BaseID]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-184
    Top =533
    Right =1332
    Bottom =1371
    Left =-1
    Top =-1
    Right =1484
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
