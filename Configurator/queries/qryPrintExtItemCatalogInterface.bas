Operation =1
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"Customer Part Number\" Or (tblItemCata"
    "logInterface.ElementName)=\"Customer Version\" Or (tblItemCatalogInterface.Eleme"
    "ntName)=\"Design Code\" Or (tblItemCatalogInterface.ElementName)=\"Fiber Type\")"
    ")"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Expression ="tblItemCatalogInterface.*"
End
Begin OrderBy
    Expression ="tblItemCatalogInterface.ElementName"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd06f4448b11b1742af42b76671815f3e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementName"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =342
        Bottom =109
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
End
