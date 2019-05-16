Operation =4
Option =0
Where ="(((AFLPRD_INVSysItemCost_CAB.Item_Cost)<>0))"
Begin InputTables
    Name ="AFLPRD_INVSysItemCost_CAB"
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Name ="tbl_Fibers.Cost"
    Expression ="[Item_Cost]"
    Name ="tbl_Fibers.Description_Oracle"
    Expression ="[Description]"
End
Begin Joins
    LeftTable ="AFLPRD_INVSysItemCost_CAB"
    RightTable ="tbl_Fibers"
    Expression ="AFLPRD_INVSysItemCost_CAB.ItemNumber = tbl_Fibers.Fiber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x776e72712dc8b94fa5e5ffb1039285ee
End
Begin
    Begin
        dbText "Name" ="tbl_Fibers.Description_Oracle"
        dbInteger "ColumnWidth" ="8925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Cost"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemCost_CAB.Item_Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =69
    Top =125
    Right =1178
    Bottom =538
    Left =-1
    Top =-1
    Right =1077
    Bottom =148
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =286
        Bottom =199
        Top =0
        Name ="AFLPRD_INVSysItemCost_CAB"
        Name =""
    End
    Begin
        Left =413
        Top =13
        Right =608
        Bottom =191
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
