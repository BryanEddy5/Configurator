Operation =1
Option =16
RowCount ="1"
Where ="(((qryfrmCreateItemTBType.TBType)=Mid(tblNewPrductConstruction.[Item No],12,1)) "
    "And ((qryfrmCreateItemTBType.Family)=Left(tblNewPrductConstruction.[Item No],2))"
    ")"
Begin InputTables
    Name ="qryfrmCreateItemTBTypeMax"
    Name ="qryfrmCreateItemTBType"
    Name ="tblNewPrductConstruction"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemTBType.TBType"
    Expression ="qryfrmCreateItemTBType.Family"
    Expression ="qryfrmCreateItemTBType.[TB Material]"
    Expression ="qryfrmCreateItemTBType.[TB Chips Type]"
    Alias ="EZStrip"
    Expression ="qryfrmCreateItemTBType.[EZ Strip]"
    Expression ="qryfrmCreateItemTBType.[TB Nominal OD]"
    Expression ="qryfrmCreateItemTBType.[CountOfTB Material]"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblNewPrductConstruction.[Jacket Material]"
End
Begin Joins
    LeftTable ="qryfrmCreateItemTBTypeMax"
    RightTable ="qryfrmCreateItemTBType"
    Expression ="qryfrmCreateItemTBTypeMax.[MaxOfCountOfTB Material] = qryfrmCreateItemTBType.[Co"
        "untOfTB Material]"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax"
    RightTable ="qryfrmCreateItemTBType"
    Expression ="qryfrmCreateItemTBTypeMax.[JacketMaterial] = qryfrmCreateItemTBType.[JacketMater"
        "ial]"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax"
    RightTable ="qryfrmCreateItemTBType"
    Expression ="qryfrmCreateItemTBTypeMax.Family = qryfrmCreateItemTBType.Family"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax"
    RightTable ="qryfrmCreateItemTBType"
    Expression ="qryfrmCreateItemTBTypeMax.TBType = qryfrmCreateItemTBType.TBType"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblNewPrductConstruction"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblNewPrductConstruction.Base = tblCableConstructionReferences.Base"
    Flag =1
End
Begin OrderBy
    Expression ="qryfrmCreateItemTBType.[CountOfTB Material]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfc2b692aca4c2749b09b83f1ef6d6eb3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.[TB Material]"
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EZStrip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xae2d7058a345bf4a878ca20db267da80
        End
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.[TB Chips Type]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType.[CountOfTB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-263
    Top =207
    Right =1273
    Bottom =803
    Left =-1
    Top =-1
    Right =1504
    Bottom =159
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1162
        Top =13
        Right =1488
        Bottom =209
        Top =0
        Name ="qryfrmCreateItemTBTypeMax"
        Name =""
    End
    Begin
        Left =779
        Top =2
        Right =1044
        Bottom =219
        Top =0
        Name ="qryfrmCreateItemTBType"
        Name =""
    End
    Begin
        Left =55
        Top =9
        Right =286
        Bottom =187
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
    Begin
        Left =308
        Top =29
        Right =452
        Bottom =173
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =500
        Top =39
        Right =741
        Bottom =194
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
