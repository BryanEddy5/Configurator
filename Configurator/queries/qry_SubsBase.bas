Operation =1
Option =0
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="OracleBase#"
    Expression ="Left([New Oracle Part #],7)"
    Alias ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2bf9eb41db91d64ead59dc7327463698
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="OracleBase#"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf6b880dec00abb40af15f54e19d4fad1
        End
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9cb57139b6ba04e831d6a5caf4e1ba5
        End
    End
End
Begin
    State =0
    Left =61
    Top =17
    Right =1898
    Bottom =341
    Left =-1
    Top =-1
    Right =1805
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
