Operation =1
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="Basic Product Construction"
    Name ="tblCableConstructionReferences"
    Name ="vSinglePassComponentParent"
    Name ="tblCableTightBufferReference"
End
Begin OutputColumns
    Expression ="vSinglePassComponentParent.Component"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Expression ="tblCableConstructionReferences.TBType"
    Expression ="vSinglePassComponentParent.TBLetter"
    Alias ="MaxOfFiberCount"
    Expression ="Max(vSinglePassComponentParent.FiberCount)"
    Expression ="tblCableTightBufferReference.TBTypeID"
    Expression ="[Basic Product Construction].Customer"
    Alias ="CableFiberCount"
    Expression ="tblCableConstructionReferences.FiberCount"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="Basic Product Construction"
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber = [Basic Product Construction]."
        "[New Oracle Part #]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vSinglePassComponentParent"
    Expression ="[Basic Product Construction].[EZ Strip] = vSinglePassComponentParent.[EZ Strip]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vSinglePassComponentParent"
    Expression ="[Basic Product Construction].[Fiber Oracle item] = vSinglePassComponentParent.Fi"
        "ber"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vSinglePassComponentParent"
    Expression ="[Basic Product Construction].[TB Material] = vSinglePassComponentParent.[TB Mate"
        "rial]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vSinglePassComponentParent"
    Expression ="[Basic Product Construction].[TB Nominal OD] = vSinglePassComponentParent.[TB No"
        "minal OD]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vSinglePassComponentParent"
    Expression ="[Basic Product Construction].RibbonHighCure = vSinglePassComponentParent.RibbonH"
        "ighCure"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableConstructionReferences.TBType = tblCableTightBufferReference.TBType"
    Flag =1
End
Begin Groups
    Expression ="vSinglePassComponentParent.Component"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
    Expression ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[EZ Strip]"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.TBType"
    GroupLevel =0
    Expression ="vSinglePassComponentParent.TBLetter"
    GroupLevel =0
    Expression ="tblCableTightBufferReference.TBTypeID"
    GroupLevel =0
    Expression ="[Basic Product Construction].Customer"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.FiberCount"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x50bd364b129afd43adc67a1daeb92512
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableFiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vSinglePassComponentParent.Component"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vSinglePassComponentParent.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfFiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x85a117dc4cccb0479c4bdceb9f8c88bd
        End
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.TBTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =86
    Right =1187
    Bottom =715
    Left =-1
    Top =-1
    Right =1077
    Bottom =308
    Left =384
    Top =0
    ColumnsShown =543
    Begin
        Left =-354
        Top =30
        Right =-64
        Bottom =226
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =-6
        Top =8
        Right =432
        Bottom =286
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =517
        Top =29
        Right =753
        Bottom =272
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =834
        Top =15
        Right =1037
        Bottom =267
        Top =0
        Name ="vSinglePassComponentParent"
        Name =""
    End
    Begin
        Left =1092
        Top =28
        Right =1348
        Bottom =281
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End
