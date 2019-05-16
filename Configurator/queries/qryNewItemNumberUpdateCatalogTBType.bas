Operation =4
Option =8
Where ="(((tblItemCatalogInterface.ElementName) Like \"*tb type*\") AND ((vCatalogTightB"
    "ufferType.[New Oracle Part #])=getItemNumberTightbuffer([ItemNumber])))"
Begin InputTables
    Name ="tblItemCatalogInterface"
    Name ="vCatalogTightBufferType"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="[OracleTBDesc]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x23abc53271d0e84c9779b61f2e414aed
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="4125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCatalogTightBufferType.OracleTBDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemCatalogInterface.itemnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCatalogTightBufferType.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =361
    Top =174
    Right =1470
    Bottom =678
    Left =-1
    Top =-1
    Right =1077
    Bottom =207
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =584
        Top =18
        Right =1047
        Bottom =257
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
    Begin
        Left =144
        Top =12
        Right =469
        Bottom =208
        Top =0
        Name ="vCatalogTightBufferType"
        Name =""
    End
End
