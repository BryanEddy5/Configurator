Operation =1
Option =0
Where ="(((qryItemNoSearch.[Item No])=[Forms]![frmCreateItem]![ItemNo]))"
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.Oracle"
    Expression ="qryItemNoSearch.[Item No]"
    Alias ="ExistingItemNo"
    Expression ="IIf([Item No] Is Not Null,-1,0)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x69dee05b8f419a4597532e7203520554
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingItemNo"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x16fa27fa705eda448bae2c42452f00d1
        End
    End
End
Begin
    State =0
    Left =64
    Top =241
    Right =1849
    Bottom =581
    Left =-1
    Top =-1
    Right =1753
    Bottom =109
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =340
        Bottom =124
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End
