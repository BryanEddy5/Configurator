Operation =4
Option =0
Begin InputTables
    Name ="Basic Product Construction"
    Name ="AFLPRD_INVSysItemCost_CAB"
End
Begin OutputColumns
    Name ="[Basic Product Construction].OracleStatus"
    Expression ="[Item_Status]"
End
Begin Joins
    LeftTable ="AFLPRD_INVSysItemCost_CAB"
    RightTable ="Basic Product Construction"
    Expression ="AFLPRD_INVSysItemCost_CAB.ItemNumber = [Basic Product Construction].[New Oracle "
        "Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa2cbb9011dbfce49b6a461f18629d02b
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =455
    Top =109
    Right =1005
    Bottom =738
    Left =-1
    Top =-1
    Right =518
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =303
        Top =29
        Right =447
        Bottom =173
        Top =0
        Name ="AFLPRD_INVSysItemCost_CAB"
        Name =""
    End
End
