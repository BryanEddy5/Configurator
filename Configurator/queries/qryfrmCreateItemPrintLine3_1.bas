Operation =1
Option =0
Begin InputTables
    Name ="qryfrmCreateItemPrintLine3"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryfrmCreateItemPrintLine3.PrintLine3"
    Alias ="CountOfPrintLine3"
    Expression ="Count(qryfrmCreateItemPrintLine3.PrintLine3)"
End
Begin Groups
    Expression ="qryfrmCreateItemPrintLine3.PrintLine3"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4c44d083325f7a4fb30e36dc873a0ee2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1f6f81e2c9bd5745bfb467746a8ba85d
        End
    End
    Begin
        dbText "Name" ="CountOfPrintLine3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x70b699e45126994b9d08db16027a9681
        End
    End
End
Begin
    State =0
    Left =104
    Top =98
    Right =1578
    Bottom =422
    Left =-1
    Top =-1
    Right =1442
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =175
        Bottom =94
        Top =0
        Name ="qryfrmCreateItemPrintLine3"
        Name =""
    End
End
