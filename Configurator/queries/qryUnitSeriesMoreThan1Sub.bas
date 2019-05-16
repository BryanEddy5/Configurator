Operation =1
Option =0
Having ="(((Count([Unit Series].[Unit Type]))<2) AND (([Unit Series].[Unit Type]) Not Lik"
    "e \"*@*\") AND ((Mid([Unit Color Series],3,3))<>Mid([Unit Type],3,3)))"
Begin InputTables
    Name ="Unit Series"
End
Begin OutputColumns
    Expression ="[Unit Series].[Unit Color Series]"
    Alias ="CountOfUnit Type"
    Expression ="Count([Unit Series].[Unit Type])"
    Expression ="[Unit Series].[Unit Type]"
    Alias ="Expr1"
    Expression ="Mid([Unit Color Series],3,3)"
End
Begin Groups
    Expression ="[Unit Series].[Unit Color Series]"
    GroupLevel =0
    Expression ="[Unit Series].[Unit Type]"
    GroupLevel =0
    Expression ="Mid([Unit Color Series],3,3)"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="((Query1.[Unit Color Series]=\"KP012666-PVDF\"))"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc2be6b8ba20de84ba2d03e893261fdcb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfUnit Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33b130ba0822564cb74da2a1b3fb1b1c
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6a02708aac19748a09a3b2d8e2034a5
        End
    End
End
Begin
    State =0
    Left =61
    Top =391
    Right =1170
    Bottom =949
    Left =-1
    Top =-1
    Right =1077
    Bottom =327
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =324
        Bottom =199
        Top =0
        Name ="Unit Series"
        Name =""
    End
End
