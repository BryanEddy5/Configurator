Operation =1
Option =0
Begin InputTables
    Name ="tblFiberSpecs"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblFiberSpecs.FiberType"
    Alias ="MaxOfFiberTypeCount"
    Expression ="Max(tblFiberSpecs.FiberTypeCount)"
End
Begin Groups
    Expression ="tblFiberSpecs.FiberType"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2fe07d1be39f0f499c324464f33e854b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MaxOfFiberTypeCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8eb0d4b6a5bcf24f9b02682856afaf78
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa23c1b1665d1eb409c82c74f8cb5a7cf
        End
    End
End
Begin
    State =0
    Left =190
    Top =130
    Right =1898
    Bottom =588
    Left =-1
    Top =-1
    Right =1676
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =235
        Bottom =184
        Top =0
        Name ="tblFiberSpecs"
        Name =""
    End
End
