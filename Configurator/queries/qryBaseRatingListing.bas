Operation =1
Option =2
Having ="(((qryBaseListing.Base) Not Like \"*ia*\") AND ((qryBaseListing.[Listing Company"
    "]) Not Like \"*none*\") AND ((qryItemNoJoin.Active)=True) AND ((qryBasePrintLine"
    "3.PrintLine3) Is Not Null And (qryBasePrintLine3.PrintLine3) Like \"*of*\") AND "
    "((tblBurnListing.Base) Is Null))"
Begin InputTables
    Name ="qryBaseListing"
    Name ="qryBasePrintLine3"
    Name ="tblBurnListing"
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Alias ="Rating"
    Expression ="IIf([PrintLine3] Like \"*ls*\",Mid([PrintLine3],InStr([PrintLine3],\"OF\"),7),Mi"
        "d([PrintLine3],InStr([PrintLine3],\"OF\"),4))"
    Expression ="qryBaseListing.Base"
    Expression ="qryBaseListing.[Listing Company]"
    Expression ="qryItemNoJoin.Active"
End
Begin Joins
    LeftTable ="qryBaseListing"
    RightTable ="qryBasePrintLine3"
    Expression ="qryBaseListing.Base = qryBasePrintLine3.Base"
    Flag =1
    LeftTable ="qryBasePrintLine3"
    RightTable ="tblBurnListing"
    Expression ="qryBasePrintLine3.Base = tblBurnListing.Base"
    Flag =2
    LeftTable ="qryBasePrintLine3"
    RightTable ="qryItemNoJoin"
    Expression ="qryBasePrintLine3.Base = qryItemNoJoin.Base"
    Flag =1
End
Begin Groups
    Expression ="IIf([PrintLine3] Like \"*ls*\",Mid([PrintLine3],InStr([PrintLine3],\"OF\"),7),Mi"
        "d([PrintLine3],InStr([PrintLine3],\"OF\"),4))"
    GroupLevel =0
    Expression ="qryBaseListing.Base"
    GroupLevel =0
    Expression ="qryBaseListing.[Listing Company]"
    GroupLevel =0
    Expression ="qryItemNoJoin.Active"
    GroupLevel =0
    Expression ="qryBasePrintLine3.PrintLine3"
    GroupLevel =0
    Expression ="tblBurnListing.Base"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe7d83608348c5544b1335dbb664cdf14
End
dbMemo "OrderBy" ="qryBaseRatingListing.Base"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryBaseListing.Base"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rating"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf102ea8cc8c82342bd977059b6fe86cb
        End
    End
End
Begin
    State =0
    Left =143
    Top =93
    Right =1869
    Bottom =558
    Left =-1
    Top =-1
    Right =1694
    Bottom =234
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="qryBaseListing"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =295
        Bottom =139
        Top =0
        Name ="qryBasePrintLine3"
        Name =""
    End
    Begin
        Left =357
        Top =10
        Right =620
        Bottom =173
        Top =0
        Name ="tblBurnListing"
        Name =""
    End
    Begin
        Left =792
        Top =6
        Right =1008
        Bottom =169
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End
