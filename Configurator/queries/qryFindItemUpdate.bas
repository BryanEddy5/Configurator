Operation =4
Option =0
Begin InputTables
    Name ="qryFindItemCopy"
End
Begin OutputColumns
    Name ="qryFindItemCopy.[New Oracle Part #]"
    Expression ="[OracleItem#]"
    Name ="qryFindItemCopy.[Item No]"
    Expression ="[ItemNo]"
    Name ="qryFindItemCopy.Customer"
    Expression ="[Customer]"
    Name ="qryFindItemCopy.[Customer Part#]"
    Expression ="[PID]"
    Name ="qryFindItemCopy.[Revision Date]"
    Expression ="Date()"
    Name ="qryFindItemCopy.[Revision Letter]"
    Expression ="\"A\""
    Name ="qryFindItemCopy.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & [INITIALS]"
    Name ="qryFindItemCopy.OracleStatus"
    Expression ="\"Active\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x79f019e23c802d4fb933ab8867b939bd
End
Begin
End
Begin
    State =0
    Left =18
    Top =191
    Right =1898
    Bottom =907
    Left =-1
    Top =-1
    Right =1848
    Bottom =485
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =538
        Bottom =454
        Top =0
        Name ="qryFindItemCopy"
        Name =""
    End
End
