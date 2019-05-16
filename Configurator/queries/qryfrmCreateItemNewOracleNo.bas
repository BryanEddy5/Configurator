Operation =1
Option =0
Where ="(((Mid([Oracle],2,1))=Left(Forms!frmMainMenu!ItemNo,1)) And ((qryFiberOracle.Ora"
    "cle) Not Like \"W*\" And (qryFiberOracle.Oracle) Like \"P*\"))"
Begin InputTables
    Name ="qryFiberOracle"
End
Begin OutputColumns
    Alias ="ItemLetter"
    Expression ="Mid([Oracle],2,1)"
    Expression ="qryFiberOracle.Oracle"
    Alias ="NewOracle"
    Expression ="Left([Oracle],2) & Format(Val(Mid([Oracle],3,5)+8),\"00000\")"
End
Begin OrderBy
    Expression ="qryFiberOracle.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x133792fb48a9e647993b7ba033d719e5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ItemLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf038d2dd744c094c9a42feaeffb26d96
        End
    End
    Begin
        dbText "Name" ="NewOracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9e8b29f498f6b1468aa3d01b01d973c1
        End
    End
End
Begin
    State =0
    Left =340
    Top =47
    Right =1814
    Bottom =688
    Left =-1
    Top =-1
    Right =1442
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
End
