Operation =1
Option =2
Where ="(((tblUnitSeries.[Unit Type]) Not Like \"*@*\") AND (([Basic Product Constructio"
    "n].Active)<>0))"
Begin InputTables
    Name ="qryPremiseItemDetails"
    Name ="tblUnitSeries"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="tblUnitSeries.[Unit Type]"
End
Begin Joins
    LeftTable ="qryPremiseItemDetails"
    RightTable ="tblUnitSeries"
    Expression ="qryPremiseItemDetails.UnitSeries = tblUnitSeries.[Unit Color Series]"
    Flag =1
    LeftTable ="tblUnitSeries"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeries.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x15fc40b8d7b2994fb22f4cbfb96cab6f
End
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =1127
    Bottom =478
    Left =-1
    Top =-1
    Right =1077
    Bottom =216
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =155
        Top =14
        Right =370
        Bottom =252
        Top =0
        Name ="qryPremiseItemDetails"
        Name =""
    End
    Begin
        Left =443
        Top =43
        Right =750
        Bottom =221
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =859
        Top =44
        Right =1029
        Bottom =192
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
