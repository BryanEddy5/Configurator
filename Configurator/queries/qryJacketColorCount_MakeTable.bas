Operation =1
Option =0
Having ="((([qryJacketColorCount].[Color]) Not Like \"*S1*\"))"
Begin InputTables
    Name ="qryJacketColorCount"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryJacketColorCount.Color"
    Alias ="Expr2"
    Expression ="qryJacketColorCount.Jacket"
    Alias ="Expr3"
    Expression ="qryJacketColorCount.ColorLetter"
    Alias ="MaxOfCount"
    Expression ="Max(qryJacketColorCount.Count)"
End
Begin OrderBy
    Expression ="qryJacketColorCount.Jacket"
    Flag =0
    Expression ="qryJacketColorCount.ColorLetter"
    Flag =0
End
Begin Groups
    Expression ="qryJacketColorCount.Color"
    GroupLevel =0
    Expression ="qryJacketColorCount.Jacket"
    GroupLevel =0
    Expression ="qryJacketColorCount.ColorLetter"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x21d5109da131224e825e58e6e2ddded0
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryJacketColorCount.Color"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColorCount.Jacket"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfCount"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22808a4ba4b5984e832b3d3f226ba22a
        End
    End
    Begin
        dbText "Name" ="qryJacketColorCount.ColorLetter"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x27fd96a69944454b91e23c6eed86c04c
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x918364f32301de4c9d0c5512eda9bd5d
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf95037e1fd0aa844945aab9565288fe5
        End
    End
End
Begin
    State =0
    Left =143
    Top =146
    Right =1435
    Bottom =830
    Left =-1
    Top =-1
    Right =1260
    Bottom =350
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =293
        Top =67
        Right =560
        Bottom =319
        Top =0
        Name ="qryJacketColorCount"
        Name =""
    End
End
