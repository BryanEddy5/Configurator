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
            0x662df329f3216147bfafbf5f5dda41ef
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
            0x9f64f5a69b75c247b25b5f807c4dc43d
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x65f5e9bcf664004e8f2d140d4a1e3f37
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb20eb6bc26686b44977f5014a877a824
        End
    End
End
Begin
    State =0
    Left =124
    Top =129
    Right =1416
    Bottom =813
    Left =-1
    Top =-1
    Right =1260
    Bottom =350
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =327
        Top =75
        Right =594
        Bottom =327
        Top =0
        Name ="qryJacketColorCount"
        Name =""
    End
End
