Operation =1
Option =16
RowCount ="1,"
Having ="(((qryJacketColor.ColorLetter)=vColorLetter()) AND ((qryJacketColor.ColorChip) N"
    "ot Like \"*s1*\")) OR (((qryJacketColor.ColorLetter)=vColorLetter()) AND ((qryJa"
    "cketColor.ColorChip) Not Like \"*str*\"))"
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
    0x65c59c8a23736544857b78c6c20d847b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_Jackets_Subs.Jacket_Color_letter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Jackets_Subs.[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.JacketBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.Jacket_Color_letter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Jackets_Subs.[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfColor Chip ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="qryJacketColor.[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Jackets_Subs.JacketBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.ColorLetter"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =66
    Top =202
    Right =1507
    Bottom =693
    Left =-1
    Top =-1
    Right =1409
    Bottom =217
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
        Left =760
        Top =33
        Right =1002
        Bottom =206
        Top =0
        Name ="qryJacketColor"
        Name =""
    End
End
