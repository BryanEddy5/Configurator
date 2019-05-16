Operation =1
Option =0
Begin InputTables
    Name ="qryBasePrintLine3"
End
Begin OutputColumns
    Expression ="qryBasePrintLine3.Base"
    Alias ="PrintLine_3"
    Expression ="Replace([PrintLine3],\"E121250 \",\"\")"
    Expression ="qryBasePrintLine3.Customer"
    Expression ="qryBasePrintLine3.[Listing Company]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x71d7c406e96f5147bed99e2091141da4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="PrintLine_3"
        dbInteger "ColumnWidth" ="5220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9bddf61c7d9da04fb0ca8b51a5cdd595
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x274aef832850d448a5dd0d780b825e6c
        End
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x807e26c7a760904e927a3c7fff684ab9
        End
    End
    Begin
        dbText "Name" ="Listing Company"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1c6ddc0dbe3c894a9189b2a75e9149b4
        End
    End
End
Begin
    State =0
    Left =125
    Top =300
    Right =1833
    Bottom =973
    Left =-1
    Top =-1
    Right =1676
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="qryBasePrintLine3"
        Name =""
    End
End
