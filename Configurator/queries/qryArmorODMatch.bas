Operation =1
Option =0
Having ="(((qryBaseNomOD.Base)=[Forms]![frmFiberMatchExisting]![Base]))"
Begin InputTables
    Name ="qryBaseNomOD"
End
Begin OutputColumns
    Expression ="qryBaseNomOD.Base"
    Alias ="BaseOD"
    Expression ="CVar([OD])"
    Alias ="CountOfOD"
    Expression ="Count(qryBaseNomOD.OD)"
End
Begin OrderBy
    Expression ="CVar([OD])"
    Flag =1
End
Begin Groups
    Expression ="qryBaseNomOD.Base"
    GroupLevel =0
    Expression ="CVar([OD])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9981fc577b9b9f40b1b7710ebb3779ff
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
        dbText "Name" ="BaseOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6fc4ef70dcc6b64a934f53811a14c748
        End
    End
End
Begin
    State =0
    Left =132
    Top =182
    Right =1797
    Bottom =536
    Left =-1
    Top =-1
    Right =1633
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =238
        Bottom =109
        Top =0
        Name ="qryBaseNomOD"
        Name =""
    End
End
