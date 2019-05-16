Operation =1
Option =0
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="CableRating"
    Expression ="CableRatingPrint([Print Line 2] & \" \" & [Print Line 3] & \" \",[Item No])"
    Expression ="[Basic Product Construction].ENumber"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Customer"
    Expression ="[Basic Product Construction].[Customer Part#]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x7870f1b0cc21974d90b9033dd1a334ae
End
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CableRating"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x461cd07719260341b8d143282cdc8b1a
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =432
        Top =12
        Right =795
        Bottom =266
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
