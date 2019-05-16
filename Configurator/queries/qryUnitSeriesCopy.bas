Operation =1
Option =16
RowCount ="1,"
Having ="(((tblUnitSeriesCopy.[TB Color Series]) Not Like \"*FILLER*\"))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
End
Begin OutputColumns
    Expression ="tblUnitSeriesCopy.[TB Color Series]"
    Expression ="tblUnitSeriesCopy.[Unit Type]"
    Alias ="CountOfUnit Type"
    Expression ="Count(tblUnitSeriesCopy.[Unit Type])"
End
Begin OrderBy
    Expression ="Count(tblUnitSeriesCopy.[Unit Type])"
    Flag =1
End
Begin Groups
    Expression ="tblUnitSeriesCopy.[TB Color Series]"
    GroupLevel =0
    Expression ="tblUnitSeriesCopy.[Unit Type]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4fa6ae2f873ad441a64956b4331fb9fa
End
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfUnit Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa7d358378076d6408439e734d0eb9a2b
        End
    End
End
Begin
    State =0
    Left =34
    Top =108
    Right =1143
    Bottom =652
    Left =-1
    Top =-1
    Right =1077
    Bottom =202
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =331
        Bottom =199
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
End
