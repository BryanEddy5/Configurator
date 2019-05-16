Operation =1
Option =0
Where ="(((qryItemNoSearch.Oracle)=[Forms]![frmCreateItem]![NewOracle#]))"
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.Oracle"
    Expression ="qryItemNoSearch.[Item No]"
    Expression ="qryItemNoSearch.Base"
    Expression ="qryItemNoSearch.Label"
    Alias ="Expr1"
    Expression ="qryItemNoSearch.Location"
    Expression ="qryItemNoSearch.Active"
    Expression ="qryItemNoSearch.Reason"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfe5282ba9e140d4caae363d4cdee0e06
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x259b001e499acb479cddedc1a6895f45
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb568e3ec2e8f534e8a548119e2beb99a
        End
    End
End
Begin
    State =0
    Left =104
    Top =72
    Right =1396
    Bottom =492
    Left =-1
    Top =-1
    Right =1260
    Bottom =189
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =121
        Top =19
        Right =353
        Bottom =182
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End
