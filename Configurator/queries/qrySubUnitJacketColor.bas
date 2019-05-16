Operation =1
Option =16
RowCount ="1,"
Having ="(((qryJacketColor.ColorLetter)=[Forms]![frmCreateItem]![SubColor]) AND ((qryJack"
    "etColor.ColorChip) Not Like \"*s1*\" And (qryJacketColor.ColorChip) Not Like \"*"
    "STR*\"))"
Begin InputTables
    Name ="tblNewPrductConstruction"
    Name ="qryJacketColor"
End
Begin OutputColumns
    Expression ="qryJacketColor.ColorLetter"
    Expression ="qryJacketColor.Color"
    Expression ="qryJacketColor.ColorChip"
    Alias ="CountOfColorChip"
    Expression ="Count(qryJacketColor.ColorChip)"
End
Begin Joins
    LeftTable ="tblNewPrductConstruction"
    RightTable ="qryJacketColor"
    Expression ="tblNewPrductConstruction.[Jacket Material] = qryJacketColor.Jacket"
    Flag =1
End
Begin OrderBy
    Expression ="Count(qryJacketColor.ColorChip)"
    Flag =1
End
Begin Groups
    Expression ="qryJacketColor.ColorLetter"
    GroupLevel =0
    Expression ="qryJacketColor.Color"
    GroupLevel =0
    Expression ="qryJacketColor.ColorChip"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x77c03fa89fad6d488d305da3fabdfde2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryJacketColor.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.ColorLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x89d79f24d1ab5245aa5a9c2e4e2bd814
        End
    End
End
Begin
    State =0
    Left =146
    Top =322
    Right =1587
    Bottom =781
    Left =-1
    Top =-1
    Right =1409
    Bottom =228
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =431
        Top =14
        Right =719
        Bottom =222
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
    Begin
        Left =998
        Top =91
        Right =1142
        Bottom =235
        Top =0
        Name ="qryJacketColor"
        Name =""
    End
End
