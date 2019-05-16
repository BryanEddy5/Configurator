Operation =1
Option =0
Where ="((([tblApplicationNames].[ApplicationID])=32))"
Begin InputTables
    Name ="tblApplicationNames"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblApplicationNames.Hyperlink"
    Alias ="Expr2"
    Expression ="tblApplicationNames.ApplicationID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xdf3a53a5838876488d56d67b093ff66f
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1e2511ecd315fe49820825c58cf20b70
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x88e88aaeeacedd46ba7883358e13540c
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1524
    Bottom =1000
    Left =-1
    Top =-1
    Right =1492
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =298
        Bottom =105
        Top =0
        Name ="tblApplicationNames"
        Name =""
    End
End
