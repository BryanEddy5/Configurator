Operation =1
Option =0
Where ="(((Mid([Oracle],2,1))=Left(Forms!frmMainMenu!ItemNo,1)) And ((qryItemNoJoin_LIVE"
    ".Oracle) Not Like \"W*\" And (qryItemNoJoin_LIVE.Oracle) Like \"P*\"))"
Begin InputTables
    Name ="qryItemNoJoin_LIVE"
End
Begin OutputColumns
    Alias ="ItemLetter"
    Expression ="Mid([Oracle],2,1)"
    Alias ="Expr1"
    Expression ="qryItemNoJoin_LIVE.Oracle"
    Alias ="NewOracle"
    Expression ="Left([Oracle],2) & Format(Val(Mid([Oracle],3,5)+8),\"00000\")"
End
Begin OrderBy
    Expression ="qryItemNoJoin_LIVE.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4ef034a07693ea4393cb278c79082f40
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ItemLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x215f241b3c398e4e86f49b9371657cd9
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x98235e3c062daa4586146dc6265ee579
        End
    End
    Begin
        dbText "Name" ="NewOracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x174e4ce412714e41805e5528ea37c63b
        End
    End
End
Begin
    State =0
    Left =286
    Top =47
    Right =1760
    Bottom =371
    Left =-1
    Top =-1
    Right =1442
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =380
        Top =6
        Right =476
        Bottom =124
        Top =0
        Name ="qryItemNoJoin_LIVE"
        Name =""
    End
End
