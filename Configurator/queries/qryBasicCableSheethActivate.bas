Operation =4
Option =0
Where ="((([Basic (Cable/Sheath)].[New Oracle Part Number])=[tblNewBasicCableHold].[New "
    "Oracle Part Number]))"
Begin InputTables
    Name ="tblNewBasicCableHold"
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="[Basic (Cable/Sheath)].Active"
    Expression ="-1"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x5fec0e27208e6e49ae24b2928bb002cd
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaa1eea11cec82a47950f8a4a0465dc8d
        End
    End
End
Begin
    State =0
    Left =104
    Top =72
    Right =1898
    Bottom =636
    Left =-1
    Top =-1
    Right =1762
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =336
        Bottom =169
        Top =0
        Name ="tblNewBasicCableHold"
        Name =""
    End
    Begin
        Left =374
        Top =6
        Right =629
        Bottom =187
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End
