Operation =1
Option =0
Begin InputTables
    Name ="tblDesignCodeJacketColor"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Alias ="Expr2"
    Expression ="tblDesignCodeJacketColor.Color"
End
Begin OrderBy
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0eae4a378338b047a531579692596e65
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22bb2a10f03fb94fbba7a16596b16f9b
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9b8ed4c0eb7ff94f86ef99d0e94cd5a2
        End
    End
End
Begin
    State =0
    Left =190
    Top =156
    Right =1802
    Bottom =480
    Left =-1
    Top =-1
    Right =1580
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="tblDesignCodeJacketColor"
        Name =""
    End
End
