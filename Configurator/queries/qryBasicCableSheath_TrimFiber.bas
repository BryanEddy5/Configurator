Operation =1
Option =0
Begin InputTables
    Name ="BasicCableSheath_LIVE"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="BasicCableSheath_LIVE.[New Oracle Part Number]"
    Alias ="Expr2"
    Expression ="BasicCableSheath_LIVE.[Item No]"
    Alias ="Expr3"
    Expression ="BasicCableSheath_LIVE.Customer"
    Alias ="FiberA"
    Expression ="Left([Fiber Type A],8)"
    Alias ="FiberB"
    Expression ="Left([Fiber Type B],8)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x540920709204be46a22c80246618853c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x96469f427ad07d49b92485693ff55479
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x386f1fd640dc264e8e61b4e7a3b0e33d
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xafdef37acac95c4f88f6a76c5213d4e5
        End
    End
    Begin
        dbText "Name" ="FiberA"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x01927e4698e1084fb3f395b215a4fc1d
        End
    End
    Begin
        dbText "Name" ="FiberB"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4991e9b9319984bbcd07fc98d13ab74
        End
    End
End
Begin
    State =0
    Left =147
    Top =101
    Right =1898
    Bottom =709
    Left =-1
    Top =-1
    Right =1719
    Bottom =219
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =337
        Bottom =233
        Top =0
        Name ="BasicCableSheath_LIVE"
        Name =""
    End
End
