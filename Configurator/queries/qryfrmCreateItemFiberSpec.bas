Operation =1
Option =0
Having ="((([Basic Product Construction].[New Oracle Part #]) Not Like \"PT*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Fiber Type]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].[1st Req Freq]"
    Expression ="[Basic Product Construction].[1st Max Atten]"
    Expression ="[Basic Product Construction].[1st Min BandW]"
    Expression ="[Basic Product Construction].[2nd Req Freq]"
    Expression ="[Basic Product Construction].[2nd Max Atten]"
    Expression ="[Basic Product Construction].[2nd Min BandW]"
    Alias ="CountOf1st Req Freq"
    Expression ="Count([Basic Product Construction].[1st Req Freq])"
End
Begin OrderBy
    Expression ="Count([Basic Product Construction].[1st Req Freq])"
    Flag =1
End
Begin Groups
    Expression ="[Basic Product Construction].[Fiber Type]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[1st Req Freq]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[1st Max Atten]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[1st Min BandW]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[2nd Req Freq]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[2nd Max Atten]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[2nd Min BandW]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8134d04ff6e2b141b93b22c9c739e75b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOf1st Req Freq"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf2b6151945f46d468c64fd4ce0197f37
        End
    End
End
Begin
    State =0
    Left =441
    Top =205
    Right =1822
    Bottom =797
    Left =-1
    Top =-1
    Right =1349
    Bottom =361
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =381
        Bottom =349
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
