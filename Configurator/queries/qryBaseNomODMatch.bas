Operation =1
Option =16
RowCount ="1"
Having ="(((qryBaseNomOD.Base)=[Forms]![frmCreateItem]![Base]))"
Begin InputTables
    Name ="qryBaseNomOD"
End
Begin OutputColumns
    Expression ="qryBaseNomOD.Base"
    Expression ="qryBaseNomOD.OD"
    Alias ="CountOfOD"
    Expression ="Count(qryBaseNomOD.OD)"
    Expression ="qryBaseNomOD.Jacket"
End
Begin OrderBy
    Expression ="Count(qryBaseNomOD.OD)"
    Flag =1
End
Begin Groups
    Expression ="qryBaseNomOD.Base"
    GroupLevel =0
    Expression ="qryBaseNomOD.OD"
    GroupLevel =0
    Expression ="qryBaseNomOD.Jacket"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x011a13e4400afc42a593163aa65f6736
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe08b4afe2e4e044092b8918e110b1878
        End
    End
    Begin
        dbText "Name" ="qryBaseNomOD.Jacket"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBaseNomOD.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBaseNomOD.OD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-88
    Top =209
    Right =1577
    Bottom =834
    Left =-1
    Top =-1
    Right =1633
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =110
        Top =18
        Right =382
        Bottom =181
        Top =0
        Name ="qryBaseNomOD"
        Name =""
    End
End
