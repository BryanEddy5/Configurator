Operation =1
Option =18
RowCount ="1"
Where ="(((qryItemNoSearchMatchBase.Active)<>0))"
Begin InputTables
    Name ="qryItemNoSearchMatchBase"
End
Begin OutputColumns
    Expression ="qryItemNoSearchMatchBase.Base"
    Expression ="qryItemNoSearchMatchBase.DesignCode"
    Expression ="qryItemNoSearchMatchBase.Active"
    Expression ="qryItemNoSearchMatchBase.TBType"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8decfff4a5c5f44db287976d7aa40705
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryItemNoSearchMatchBase.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearchMatchBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearchMatchBase.TBType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =86
    Top =142
    Right =1864
    Bottom =979
    Left =-1
    Top =-1
    Right =1746
    Bottom =169
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =435
        Top =36
        Right =626
        Bottom =179
        Top =0
        Name ="qryItemNoSearchMatchBase"
        Name =""
    End
End
