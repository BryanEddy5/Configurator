Operation =1
Option =0
Where ="(((tblPremiseLotSizes.ItemPrefix)=Left([Forms]![frmNewItemNumber]![CatalogNumber"
    "],1)))"
Begin InputTables
    Name ="tblPremiseLotSizes"
End
Begin OutputColumns
    Expression ="tblPremiseLotSizes.ItemPrefix"
    Expression ="tblPremiseLotSizes.LotSize"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6beade5c9849154980f625373c63b850
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblPremiseLotSizes.ItemPrefix"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =54
    Top =281
    Right =1163
    Bottom =910
    Left =-1
    Top =-1
    Right =1077
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =216
        Bottom =128
        Top =0
        Name ="tblPremiseLotSizes"
        Name =""
    End
End
