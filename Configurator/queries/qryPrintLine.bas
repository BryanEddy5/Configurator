Operation =1
Option =0
Begin InputTables
    Name ="qryBasePrint"
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Expression ="qryItemNoJoin.[Item No]"
    Expression ="qryBasePrint.Cust"
    Expression ="qryBasePrint.PrintLine1"
    Expression ="qryBasePrint.PrintLine2"
    Expression ="qryBasePrint.PrintLine3"
    Alias ="Print4"
    Expression ="IIf([PrintLine4] Like \"*meter*\",\"M\",\"FT\")"
    Expression ="qryItemNoJoin.Oracle"
End
Begin Joins
    LeftTable ="qryBasePrint"
    RightTable ="qryItemNoJoin"
    Expression ="qryBasePrint.Oracle = qryItemNoJoin.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc19e28302f9fab4180dacc41f5b3ef54
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryBasePrint.PrintLine3"
        dbInteger "ColumnWidth" ="5940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Print4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5b5a43089e9f7e48bbef15b512dda581
        End
    End
    Begin
        dbText "Name" ="Item No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x900e71f2163a5a4598ab228c76d8f8b3
        End
    End
    Begin
        dbText "Name" ="Cust"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7dc989a2760603478c546398cb0fb07a
        End
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x78f541bf2fd6074ebe2fd0d3458804d7
        End
    End
End
Begin
    State =0
    Left =153
    Top =154
    Right =1262
    Bottom =540
    Left =-1
    Top =-1
    Right =1077
    Bottom =155
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =276
        Bottom =154
        Top =0
        Name ="qryBasePrint"
        Name =""
    End
    Begin
        Left =314
        Top =6
        Right =515
        Bottom =154
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End
