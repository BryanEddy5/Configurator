Operation =1
Option =0
Where ="((([tblAFLPRD_INVSysItemSpec_CAB].[SpecificationElement]) Like \"*label*\"))"
Begin InputTables
    Name ="tblAFLPRD_INVSysItemSpec_CAB"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblAFLPRD_INVSysItemSpec_CAB.SpecName"
    Alias ="Expr2"
    Expression ="tblAFLPRD_INVSysItemSpec_CAB.ItemNumber"
    Alias ="Expr3"
    Expression ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
    Alias ="Expr4"
    Expression ="tblAFLPRD_INVSysItemSpec_CAB.TargetValue"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x71388d907d8d154e85ed2594c3c7a988
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemSpec_CAB.SpecificationElement"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6f3d19019e002e4fb7e245d5a4085506
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc7d09985adaeb94aa435f0121f46f7ad
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeb58afe6836e0a4e8fd67c6df09d808d
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9929652874232c4e89f7885220ecf46d
        End
    End
End
Begin
    State =0
    Left =49
    Top =115
    Right =1158
    Bottom =744
    Left =-1
    Top =-1
    Right =1077
    Bottom =156
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =387
        Bottom =174
        Top =0
        Name ="tblAFLPRD_INVSysItemSpec_CAB"
        Name =""
    End
End
