Operation =1
Option =18
RowCount ="1,"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="qrytblSubUnitHoldSelect"
End
Begin OutputColumns
    Alias ="Exists"
    Expression ="IIf([qrytblSubUnitHoldSelect]![New Oracle Part #] Is Not Null,-1,0)"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="qrytblSubUnitHoldSelect"
    Expression ="[Basic Product Construction].[New Oracle Part #] = qrytblSubUnitHoldSelect.[New "
        "Oracle Part #]"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([qrytblSubUnitHoldSelect]![New Oracle Part #] Is Not Null,-1,0)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2cdacb4c3561ff48affe53e375fe584b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Exists"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfcef11461531e740b3579c43f41f2f34
        End
    End
End
Begin
    State =0
    Left =29
    Top =144
    Right =1138
    Bottom =599
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
        Left =502
        Top =38
        Right =807
        Bottom =186
        Top =0
        Name ="qrytblSubUnitHoldSelect"
        Name =""
    End
End
