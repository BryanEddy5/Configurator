Operation =1
Option =0
Having ="(((tblItemLoadTempBOM.ComponentItemNumber) Like \"rbn*\" Or (tblItemLoadTempBOM."
    "ComponentItemNumber) Like \"swr*\") And ((([Basic Product Construction].[New Ora"
    "cle Part #]) Like \"rbn*\" Or ([Basic Product Construction].[New Oracle Part #])"
    " Like \"swr*\") And ([Basic Product Construction].[New Oracle Part #]) Not Like "
    "\"*osp*\") And (([Basic Product Construction].[Fiber Oracle item])=Forms!frmPrem"
    "iseTBReplace!FiberType) And (([Basic Product Construction_2].[New Oracle Part #]"
    ")=Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber) And (([Bas"
    "ic Product Construction].FiberType2)=Forms!frmItemInterface!frmPremiseItemDetail"
    "s.Form!FiberType2) And (([Basic Product Construction].Active)<>0))"
Begin InputTables
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_2"
    Name ="tblItemLoadTempBOM"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="tblItemLoadTempBOM.ComponentItemNumber"
    Alias ="NewComponentItem"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Alias ="CopiedItemNo"
    Expression ="[Basic Product Construction_1].[Item No]"
    Alias ="NewItemNo"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction_2].[New Oracle Part #]"
    Expression ="[Basic Product Construction_2].RibbonHighCure"
    Expression ="[Basic Product Construction].FiberType2"
    Expression ="[Basic Product Construction].Active"
End
Begin Joins
    LeftTable ="tblItemLoadTempBOM"
    RightTable ="Basic Product Construction_1"
    Expression ="tblItemLoadTempBOM.ComponentItemNumber = [Basic Product Construction_1].[New Ora"
        "cle Part #]"
    Flag =1
    LeftTable ="Basic Product Construction_1"
    RightTable ="Basic Product Construction"
    Expression ="[Basic Product Construction_1].Base = [Basic Product Construction].Base"
    Flag =1
    LeftTable ="Basic Product Construction_2"
    RightTable ="Basic Product Construction"
    Expression ="[Basic Product Construction_2].RibbonHighCure = [Basic Product Construction].Rib"
        "bonHighCure"
    Flag =1
End
Begin Groups
    Expression ="tblItemLoadTempBOM.ComponentItemNumber"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    GroupLevel =0
    Expression ="[Basic Product Construction_1].[Item No]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Item No]"
    GroupLevel =0
    Expression ="[Basic Product Construction_2].[New Oracle Part #]"
    GroupLevel =0
    Expression ="[Basic Product Construction_2].RibbonHighCure"
    GroupLevel =0
    Expression ="[Basic Product Construction].FiberType2"
    GroupLevel =0
    Expression ="[Basic Product Construction].Active"
    GroupLevel =0
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
    0x02a926bf2e04f14e95a78bf599b1a1e7
End
Begin
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CopiedItemNo"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc000619ca9b3af48a9aedd50d29aff54
        End
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NewItemNo"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe25fff51f7c21b4cb24345a1c73e8970
        End
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NewComponentItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x640db18e8193434d9a9221a620009209
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction_2].RibbonHighCure"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_2].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-319
    Top =169
    Right =790
    Bottom =798
    Left =-1
    Top =-1
    Right =1077
    Bottom =322
    Left =180
    Top =0
    ColumnsShown =543
    Begin
        Left =1159
        Top =11
        Right =1446
        Bottom =268
        Top =0
        Name ="Basic Product Construction_2"
        Name =""
    End
    Begin
        Left =74
        Top =22
        Right =304
        Bottom =242
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
    Begin
        Left =339
        Top =9
        Right =579
        Bottom =248
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
    Begin
        Left =665
        Top =13
        Right =1039
        Bottom =299
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
