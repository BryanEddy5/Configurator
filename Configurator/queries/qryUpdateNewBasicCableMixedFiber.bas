Operation =4
Option =0
Begin InputTables
    Name ="tblNewBasicCableHold_New"
End
Begin OutputColumns
    Name ="tblNewBasicCableHold_New.FiberType2"
    Expression ="Forms!frmFiberSpec![Fiber Oracle item]"
    Name ="tblNewBasicCableHold_New.[1st Req Freq - B]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewBasicCableHold_New.[1st Max Atten - B]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewBasicCableHold_New.[1st Min BandW - B]"
    Expression ="IIf(Forms!frmFiberSpec![1stMinBandW]=\"\",Null,Forms!frmFiberSpec![1stMinBandW])"
    Name ="tblNewBasicCableHold_New.[2nd Req Freq - B]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewBasicCableHold_New.[2nd Max Atten - B]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewBasicCableHold_New.[2nd Min BandW - B]"
    Expression ="IIf(Forms!frmFiberSpec![2ndMinBandW]=\"\",Null,Forms!frmFiberSpec![2ndMinBandW])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd0d474f1f1596d449bb84b347eb0f661
End
Begin
    Begin
        dbText "Name" ="tblNewBasicCableHold.[Fiber Type B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[Fiber Type B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[1st Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[2nd Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[1st Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[1st Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[2nd Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[2nd Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[1st Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[1st Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold_New.[2nd Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[1st Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[2nd Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCableHold.[2nd Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =184
    Top =202
    Right =1293
    Bottom =632
    Left =-1
    Top =-1
    Right =1077
    Bottom =182
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =526
        Top =15
        Right =820
        Bottom =208
        Top =0
        Name ="tblNewBasicCableHold_New"
        Name =""
    End
End
