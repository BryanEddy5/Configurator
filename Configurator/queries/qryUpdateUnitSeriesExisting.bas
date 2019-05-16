Operation =4
Option =0
Begin InputTables
    Name ="tblNewBasicCableHold"
    Name ="tblNewBasicCable_New"
End
Begin OutputColumns
    Name ="tblNewBasicCable_New.[Unit Series]"
    Expression ="[Forms]![frmCreateitem]![UnitSeries]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xac6a70f2dfa7f44b935963791a0f1c58
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)]![New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =53
    Top =256
    Right =1162
    Bottom =885
    Left =-1
    Top =-1
    Right =1077
    Bottom =161
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =688
        Top =8
        Right =1008
        Bottom =186
        Top =0
        Name ="tblNewBasicCableHold"
        Name =""
    End
    Begin
        Left =278
        Top =16
        Right =512
        Bottom =213
        Top =0
        Name ="tblNewBasicCable_New"
        Name =""
    End
End
