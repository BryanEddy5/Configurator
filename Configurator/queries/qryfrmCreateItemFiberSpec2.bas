Operation =1
Option =18
RowCount ="1,"
Having ="(((qryfrmCreateItemFiberSpec.[Fiber Oracle item])=[Forms]![frmCreateItem]![Fiber"
    "]))"
Begin InputTables
    Name ="qryfrmCreateItemFiberSpec"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemFiberSpec.[Fiber Type]"
    Expression ="qryfrmCreateItemFiberSpec.[Fiber Oracle item]"
    Expression ="qryfrmCreateItemFiberSpec.[1st Req Freq]"
    Expression ="qryfrmCreateItemFiberSpec.[1st Max Atten]"
    Expression ="qryfrmCreateItemFiberSpec.[1st Min BandW]"
    Expression ="qryfrmCreateItemFiberSpec.[2nd Req Freq]"
    Expression ="qryfrmCreateItemFiberSpec.[2nd Max Atten]"
    Expression ="qryfrmCreateItemFiberSpec.[2nd Min BandW]"
    Alias ="CountOf1st Max Atten"
    Expression ="Count(qryfrmCreateItemFiberSpec.[1st Max Atten])"
End
Begin OrderBy
    Expression ="Count(qryfrmCreateItemFiberSpec.[1st Max Atten])"
    Flag =1
End
Begin Groups
    Expression ="qryfrmCreateItemFiberSpec.[Fiber Type]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[Fiber Oracle item]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[1st Req Freq]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[1st Max Atten]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[1st Min BandW]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[2nd Req Freq]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[2nd Max Atten]"
    GroupLevel =0
    Expression ="qryfrmCreateItemFiberSpec.[2nd Min BandW]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe90a3d791ec79848ab4c16044940f6fe
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOf1st Max Atten"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x21fceb69819f2b49af470781467dd80b
        End
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[Fiber Oracle item]"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[Fiber Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[2nd Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[2nd Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[1st Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[1st Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[1st Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemFiberSpec.[2nd Req Freq]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =143
    Top =138
    Right =1885
    Bottom =594
    Left =-1
    Top =-1
    Right =1710
    Bottom =191
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =408
        Bottom =199
        Top =0
        Name ="qryfrmCreateItemFiberSpec"
        Name =""
    End
End
