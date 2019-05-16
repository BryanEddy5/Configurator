Operation =1
Option =0
Where ="(((qryfrmCreateItemTBType_ForTB.TBType)=Mid(tblNewPrductConstruction.[Item No],1"
    "2,1)) And ((qryfrmCreateItemTBType_ForTB.Family)=Left(tblNewPrductConstruction.["
    "Item No],2)) And ((qryfrmCreateItemTBType_ForTB.[Jacket Material])=[JacketMateri"
    "al]))"
Begin InputTables
    Name ="qryfrmCreateItemTBTypeMax_ForTB"
    Name ="qryfrmCreateItemTBType_ForTB"
    Name ="tblNewPrductConstruction"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Expression ="qryfrmCreateItemTBType_ForTB.TBType"
    Expression ="qryfrmCreateItemTBType_ForTB.Family"
    Expression ="qryfrmCreateItemTBType_ForTB.[TB Material]"
    Expression ="qryfrmCreateItemTBType_ForTB.[TB Chips Type]"
    Alias ="EZStrip"
    Expression ="qryfrmCreateItemTBType_ForTB.[EZ Strip]"
    Alias ="TB Nominal OD"
    Expression ="IIf([Forms]![frmCreateItem]![NomOD] Is Null,Null,[Forms]![frmCreateItem]![NomOD]"
        ")"
    Expression ="qryfrmCreateItemTBType_ForTB.[Jacket Material]"
    Expression ="qryfrmCreateItemTBType_ForTB.[CountOfTB Material]"
End
Begin Joins
    LeftTable ="qryfrmCreateItemTBTypeMax_ForTB"
    RightTable ="qryfrmCreateItemTBType_ForTB"
    Expression ="qryfrmCreateItemTBTypeMax_ForTB.[MaxOfCountOfTB Material] = qryfrmCreateItemTBTy"
        "pe_ForTB.[CountOfTB Material]"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax_ForTB"
    RightTable ="qryfrmCreateItemTBType_ForTB"
    Expression ="qryfrmCreateItemTBTypeMax_ForTB.[Jacket Material] = qryfrmCreateItemTBType_ForTB"
        ".[Jacket Material]"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax_ForTB"
    RightTable ="qryfrmCreateItemTBType_ForTB"
    Expression ="qryfrmCreateItemTBTypeMax_ForTB.TBType = qryfrmCreateItemTBType_ForTB.TBType"
    Flag =1
    LeftTable ="qryfrmCreateItemTBTypeMax_ForTB"
    RightTable ="qryfrmCreateItemTBType_ForTB"
    Expression ="qryfrmCreateItemTBTypeMax_ForTB.Family = qryfrmCreateItemTBType_ForTB.Family"
    Flag =1
    LeftTable ="tblNewPrductConstruction"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblNewPrductConstruction.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
End
Begin OrderBy
    Expression ="qryfrmCreateItemTBType_ForTB.[CountOfTB Material]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xab861dd336939542a5a814b37fdb8a6d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.[TB Material]"
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.[CountOfTB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.[TB Chips Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TB Nominal OD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0xd28b88b4c9a6ce42923e7605b705ae5e
        End
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_ForTB.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EZStrip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x81c39c4398a3aa45a5e633589130913f
        End
    End
End
Begin
    State =0
    Left =139
    Top =57
    Right =1513
    Bottom =832
    Left =-1
    Top =-1
    Right =1342
    Bottom =170
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1041
        Top =34
        Right =1292
        Bottom =216
        Top =0
        Name ="qryfrmCreateItemTBTypeMax_ForTB"
        Name =""
    End
    Begin
        Left =744
        Top =19
        Right =964
        Bottom =197
        Top =0
        Name ="qryfrmCreateItemTBType_ForTB"
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
        Left =340
        Top =31
        Right =484
        Bottom =175
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =559
        Top =53
        Right =703
        Bottom =197
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End
