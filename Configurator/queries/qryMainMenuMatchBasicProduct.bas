Operation =1
Option =0
Begin InputTables
    Name ="qryMainMenuMatch"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
End
Begin Joins
    LeftTable ="qryMainMenuMatch"
    RightTable ="Basic Product Construction"
    Expression ="qryMainMenuMatch.Oracle = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6b90ac6bd54f5f41b8bcb13a5bad16fa
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =61
    Top =43
    Right =1353
    Bottom =412
    Left =-1
    Top =-1
    Right =1260
    Bottom =138
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="qryMainMenuMatch"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =372
        Bottom =154
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
