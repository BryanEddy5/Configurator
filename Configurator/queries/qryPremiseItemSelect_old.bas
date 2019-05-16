Operation =1
Option =0
Begin InputTables
    Name ="qryPremiseItems"
End
Begin OutputColumns
    Expression ="qryPremiseItems.OracleItemNumber"
    Expression ="qryPremiseItems.CatalogNumber"
    Expression ="qryPremiseItems.Customer"
End
Begin OrderBy
    Expression ="qryPremiseItems.OracleItemNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x97019714ba49dd4aafcfe65ed40ab647
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryPremiseItems.OracleItemNumber"
        dbMemo "Caption" ="Oracle Item Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItems.CatalogNumber"
        dbMemo "Caption" ="Catalog Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =215
    Top =295
    Right =1324
    Bottom =924
    Left =-1
    Top =-1
    Right =1077
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =572
        Bottom =128
        Top =0
        Name ="qryPremiseItems"
        Name =""
    End
End
