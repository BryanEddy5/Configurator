Operation =1
Option =0
Where ="(((qryBasePrint.Cust) Not Like \"*Beld*\") AND ((qryBasePrint.PrintLine1) Not Li"
    "ke \"*afl*\") AND ((qryBasePrint.PrintLine3) Not Like \"*115304*\" And (qryBaseP"
    "rint.PrintLine3) Not Like \"*E121250*\" And ((qryBasePrint.PrintLine3) Like \"*U"
    "L*\" Or (qryBasePrint.PrintLine3) Like \"*ETL*\")))"
Begin InputTables
    Name ="qryBasePrint"
End
Begin OutputColumns
    Expression ="qryBasePrint.Cust"
    Expression ="qryBasePrint.PrintLine1"
    Expression ="qryBasePrint.PrintLine2"
    Expression ="qryBasePrint.PrintLine3"
    Expression ="qryBasePrint.Oracle"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa33799b87f403441b4818f7f1fbfde2b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.PrintLine3"
        dbInteger "ColumnWidth" ="6150"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.Customer"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBasePrint.Cust"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =147
    Top =206
    Right =1256
    Bottom =835
    Left =-1
    Top =-1
    Right =1077
    Bottom =142
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =276
        Bottom =132
        Top =0
        Name ="qryBasePrint"
        Name =""
    End
End
