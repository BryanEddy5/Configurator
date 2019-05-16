Operation =1
Option =0
Begin InputTables
    Name ="qryBaseConstruction"
End
Begin OutputColumns
    Expression ="qryBaseConstruction.Base"
    Alias ="MaxOfCountOfBase"
    Expression ="Max(qryBaseConstruction.CountOfBase)"
End
Begin OrderBy
    Expression ="qryBaseConstruction.Base"
    Flag =0
    Expression ="Max(qryBaseConstruction.CountOfBase)"
    Flag =1
End
Begin Groups
    Expression ="qryBaseConstruction.Base"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd2db63368f3ceb45854f2540e04c0120
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MaxOfCountOfBase"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb697eefb0b6124fb3ab6cd6fd304efe
        End
    End
End
Begin
    State =0
    Left =132
    Top =227
    Right =1840
    Bottom =762
    Left =-1
    Top =-1
    Right =1676
    Bottom =190
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =399
        Bottom =219
        Top =0
        Name ="qryBaseConstruction"
        Name =""
    End
End
