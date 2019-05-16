Operation =1
Option =0
Having ="(((qryBasePrintLine3_1.[Listing Company])=Forms!frmCreateItem!ListingCompany) An"
    "d ((qryBasePrintLine3_1.Base)=Forms!frmCreateItem!Base))"
Begin InputTables
    Name ="qryBasePrintLine3_1"
End
Begin OutputColumns
    Expression ="qryBasePrintLine3_1.PrintLine_3"
    Alias ="CountOfPrintLine_3"
    Expression ="Count(qryBasePrintLine3_1.PrintLine_3)"
    Expression ="qryBasePrintLine3_1.[Listing Company]"
End
Begin OrderBy
    Expression ="Count(qryBasePrintLine3_1.PrintLine_3)"
    Flag =1
End
Begin Groups
    Expression ="qryBasePrintLine3_1.PrintLine_3"
    GroupLevel =0
    Expression ="qryBasePrintLine3_1.[Listing Company]"
    GroupLevel =0
    Expression ="qryBasePrintLine3_1.Base"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x736339f88e71c1449c95058e5aca1bb5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfPrintLine_3"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x645b2abfe789a546bd920be7109d3597
        End
    End
End
Begin
    State =0
    Left =218
    Top =313
    Right =1830
    Bottom =773
    Left =-1
    Top =-1
    Right =1580
    Bottom =229
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =158
        Top =64
        Right =416
        Bottom =182
        Top =0
        Name ="qryBasePrintLine3_1"
        Name =""
    End
End
