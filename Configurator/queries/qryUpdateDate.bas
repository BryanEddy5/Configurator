Operation =1
Option =16
RowCount ="1"
Begin InputTables
    Name ="tblUpdateDate"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblUpdateDate.UpdateDate"
End
Begin OrderBy
    Expression ="tblUpdateDate.UpdateDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe113d5aff1d185458db87733163d5eaf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x464287173814f54da18e1d42fe0a9dc5
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =69
        Top =23
        Right =265
        Bottom =92
        Top =0
        Name ="tblUpdateDate"
        Name =""
    End
End
