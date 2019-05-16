Operation =1
Option =2
Where ="((([Basic Product Construction].[UL-ETL-Section]) Is Not Null) AND (([Basic Prod"
    "uct Construction].[\"E\" Number]) Is Not Null And ([Basic Product Construction]."
    "[\"E\" Number]) Like \"E*\") AND (([Basic Product Construction].[Print Line 3]) "
    "Is Not Null) AND (([Basic Product Construction].Active)=True)) OR ((([Basic Prod"
    "uct Construction].[UL-ETL-Section]) Is Not Null) AND (([Basic Product Constructi"
    "on].[\"E\" Number]) Like \"1*\") AND (([Basic Product Construction].[Print Line "
    "3]) Is Not Null) AND (([Basic Product Construction].Active)=True))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[UL-ETL-Section]"
    Alias ="Expr1"
    Expression ="[Basic Product Construction].[\"E\" Number]"
    Expression ="[Basic Product Construction].[Listing Company]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfc3022bc44cdd045b2f5b2c8f923dd5f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Section]"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[\"E\" Number]"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 3]"
        dbInteger "ColumnWidth" ="6030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf56e3066c270a54cb729af760dd36076
        End
    End
End
Begin
    State =0
    Left =51
    Top =209
    Right =1888
    Bottom =879
    Left =-1
    Top =-1
    Right =1805
    Bottom =269
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =445
        Bottom =252
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
