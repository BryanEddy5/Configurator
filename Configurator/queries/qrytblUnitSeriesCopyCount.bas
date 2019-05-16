Operation =1
Option =0
Having ="(((tblUnitSeriesCopy.[TB Color Series]) Not Like \"*fill*\"))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
End
Begin OutputColumns
    Alias ="Count"
    Expression ="Count(tblUnitSeriesCopy.[Unit ID])"
End
Begin Groups
    Expression ="tblUnitSeriesCopy.[TB Color Series]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2dc1407c36582640b136cfbc9852cfd5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x62c6892a4a183e44ac2b1b0321b4f619
        End
    End
End
Begin
    State =0
    Left =17
    Top =239
    Right =1126
    Bottom =649
    Left =-1
    Top =-1
    Right =1077
    Bottom =162
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =262
        Bottom =184
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
End
