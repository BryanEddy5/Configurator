Operation =1
Option =18
RowCount ="1,"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="qrytblSubUnit_X_Hold"
End
Begin OutputColumns
    Alias ="Existing"
    Expression ="IIf([qrytblSubUnit_X_Hold]![New Oracle Part #] Is Not Null,-1,0)"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="qrytblSubUnit_X_Hold"
    Expression ="[Basic Product Construction].[New Oracle Part #] = qrytblSubUnit_X_Hold.[New Ora"
        "cle Part #]"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([qrytblSubUnit_X_Hold]![New Oracle Part #] Is Not Null,-1,0)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xed22007282b5eb4db71c896b4e2469a7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Existing"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7e10dce47c60f3438c6a4283638a773b
        End
    End
End
Begin
    State =0
    Left =56
    Top =66
    Right =1165
    Bottom =521
    Left =-1
    Top =-1
    Right =1077
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =299
        Bottom =214
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =617
        Top =38
        Right =939
        Bottom =201
        Top =0
        Name ="qrytblSubUnit_X_Hold"
        Name =""
    End
End
