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
    LeftTable ="tblCableConstructionReferences"
    RightTable ="vSinglePassComponentParent"
    Expression ="tblCableConstructionReferences.FibersPerBundle = vSinglePassComponentParent.Fibe"
        "rsPerBundle"
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
dbBinary "GUID" = Begin
    0x250dddcc66a9754fac89c1181e1dd975
End
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
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
            0xe23efc298087b2479fefbf5c4024ebf7
        End
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
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
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =225
    Left =97
    Top =0
    ColumnsShown =543
    Begin
        Left =14
        Top =27
        Right =304
        Bottom =223
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =329
        Top =-1
        Right =767
        Bottom =277
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =829
        Top =21
        Right =1065
        Bottom =264
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1121
        Top =15
        Right =1324
        Bottom =267
        Top =0
        Name ="vSinglePassComponentParent"
        Name =""
    End
    Begin
        Left =1379
        Top =28
        Right =1635
        Bottom =281
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End
