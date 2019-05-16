Operation =1
Option =0
Begin InputTables
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Alias ="OracleBase#"
    Expression ="Left([New Oracle Part Number],7)"
    Alias ="Expr1"
    Expression ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x390e8824b4b8144b90fd219bfa3f36cc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="OracleBase#"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45da188547aa6f4a84da46f492758885
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06de24e4c0a5ea40bad21215da88370b
        End
    End
End
Begin
    State =0
    Left =398
    Top =23
    Right =1771
    Bottom =479
    Left =-1
    Top =-1
    Right =1341
    Bottom =225
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =373
        Bottom =244
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End
