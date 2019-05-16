Operation =1
Option =2
Having ="(((qryBasePrintLine2.PrintLine2)<>\"\") And ((qryBasePrintLine2.Customer) Like L"
    "eft(Forms!frmCreateItem!Customer,10) & \"*\") And ((qryBasePrintLine2.Fiber)=For"
    "ms!frmCreateItem!Fiber))"
Begin InputTables
    Name ="qryBasePrintLine2"
End
Begin OutputColumns
    Expression ="qryBasePrintLine2.PrintLine2"
    Alias ="CountOfPrintLine2"
    Expression ="Count(qryBasePrintLine2.PrintLine2)"
End
Begin OrderBy
    Expression ="Count(qryBasePrintLine2.PrintLine2)"
    Flag =1
End
Begin Groups
    Expression ="qryBasePrintLine2.PrintLine2"
    GroupLevel =0
    Expression ="qryBasePrintLine2.Customer"
    GroupLevel =0
    Expression ="qryBasePrintLine2.Fiber"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8754b2623aa1ac4e8f433d0999b1bcbb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfPrintLine2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd521af4dfa8d124496ecbd1063a9bca6
        End
    End
    Begin
        dbText "Name" ="qryBasePrintLine2.PrintLine2"
        dbInteger "ColumnWidth" ="4890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =112
    Top =388
    Right =1724
    Bottom =851
    Left =-1
    Top =-1
    Right =1580
    Bottom =157
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =311
        Top =6
        Right =570
        Bottom =154
        Top =0
        Name ="qryBasePrintLine2"
        Name =""
    End
End
