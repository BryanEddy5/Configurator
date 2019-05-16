Operation =1
Option =0
Where ="(((qryBasePrintLine3.PrintLine3) Is Not Null))"
Begin InputTables
    Name ="qryBasePrintLine3"
    Name ="qryBaseListing"
End
Begin OutputColumns
    Alias ="Rating"
    Expression ="IIf([PrintLine3] Like \"*ls*\",Mid([PrintLine3],InStr([PrintLine3],\"OF\"),7),Mi"
        "d([PrintLine3],InStr([PrintLine3],\"OF\"),4))"
    Expression ="qryBasePrintLine3.Base"
    Expression ="qryBaseListing.[Listing Company]"
End
Begin Joins
    LeftTable ="qryBasePrintLine3"
    RightTable ="qryBaseListing"
    Expression ="qryBasePrintLine3.Base = qryBaseListing.Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe229b3ea95e821479b2ad8331091c5fc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Rating"
        dbInteger "ColumnWidth" ="3135"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ba1e0baf922ca41af13ca18322f19f7
        End
    End
End
Begin
    State =0
    Left =114
    Top =158
    Right =1223
    Bottom =658
    Left =-1
    Top =-1
    Right =1077
    Bottom =269
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =446
        Top =6
        Right =579
        Bottom =124
        Top =0
        Name ="qryBasePrintLine3"
        Name =""
    End
    Begin
        Left =663
        Top =23
        Right =759
        Bottom =111
        Top =0
        Name ="qryBaseListing"
        Name =""
    End
End
