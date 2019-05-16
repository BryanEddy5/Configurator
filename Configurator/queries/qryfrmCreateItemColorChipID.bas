Operation =1
Option =0
Having ="(((Count(qryJacketColor.ColorChip))>0) And ((qryJacketColor.Jacket)=Forms!frmCre"
    "ateItem!Jacket) And ((qryJacketColor.ColorLetter)=Right(Left(Forms!frmCreateItem"
    "!ItemNo,10),1)))"
Begin InputTables
    Name ="qryJacketColor"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryJacketColor.ColorChip"
    Alias ="CountOfColorChip"
    Expression ="Count(qryJacketColor.ColorChip)"
End
Begin OrderBy
    Expression ="Count(qryJacketColor.ColorChip)"
    Flag =1
End
Begin Groups
    Expression ="qryJacketColor.ColorChip"
    GroupLevel =0
    Expression ="qryJacketColor.Jacket"
    GroupLevel =0
    Expression ="qryJacketColor.ColorLetter"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x156ad910e39152459eea6103efcd69bd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4bf6cd1475d54b43bd2698838aa8f2c0
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa4faf260cd1c8d4b8915c1e424057a26
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =142
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =11
        Right =327
        Bottom =144
        Top =0
        Name ="qryJacketColor"
        Name =""
    End
End
