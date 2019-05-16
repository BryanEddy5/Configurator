Operation =1
Option =0
Where ="((([tblTightBuffer].[TBLetter])<>\"e\" And ([tblTightBuffer].[TBLetter])<>\"s\" "
    "And ([tblTightBuffer].[TBLetter])<>\"4\" And ([tblTightBuffer].[TBLetter])<>\"5\""
    " And ([tblTightBuffer].[TBLetter])<>\"b\" And ([tblTightBuffer].[TBLetter])<>\"0"
    "\"))"
Begin InputTables
    Name ="tblTightBuffer"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblTightBuffer.TBLetter"
    Alias ="Expr2"
    Expression ="tblTightBuffer.TBDesc"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf3ef4e37cb63a943a699f0dacbb2b7f4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblTightBuffer.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTightBuffer.TBDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc06fd31f1490104fab441cec1728fb6d
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x57907703149efe428208eb7568261642
        End
    End
End
Begin
    State =0
    Left =-227
    Top =326
    Right =1558
    Bottom =650
    Left =-1
    Top =-1
    Right =1753
    Bottom =59
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="tblTightBuffer"
        Name =""
    End
End
