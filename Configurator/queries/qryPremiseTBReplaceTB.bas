Operation =1
Option =0
Having ="(((qryPremiseTBItems.[Fiber Oracle item])=Forms!frmPremiseTBReplace!FiberType) A"
    "nd (([Basic Product Construction].[New Oracle Part #])=Forms!frmItemInterface!fr"
    "mPremiseItemDetails.Form!OracleItemNumber))"
Begin InputTables
    Name ="qryPremiseTBItems"
    Name ="qryPremiseTBItemsCopied"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="qryPremiseTBItemsCopied.ComponentItemNumber"
    Alias ="CopiedItemNo"
    Expression ="qryPremiseTBItemsCopied.[Item No]"
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    Expression ="qryPremiseTBItems.[Jacket Material]"
    Alias ="NewComponentItem"
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    Alias ="NewItemNo"
    Expression ="qryPremiseTBItems.[Item No]"
    Alias ="NewBase"
    Expression ="Left(qryPremiseTBItems.[New Oracle Part #],7)"
    Expression ="qryPremiseTBItemsCopied.[Fiber Oracle item]"
    Expression ="qryPremiseTBItemsCopied.[Jacket Material]"
    Expression ="qryPremiseTBItems.[Item No]"
    Expression ="qryPremiseTBItemsCopied.[Item No]"
    Alias ="TB"
    Expression ="Mid(qryPremiseTBItemsCopied.[Item No],12,1)"
    Alias ="Expr1"
    Expression ="Mid(qryPremiseTBItems![Item No],12,1)"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
Begin Joins
    LeftTable ="qryPremiseTBItems"
    RightTable ="qryPremiseTBItemsCopied"
    Expression ="qryPremiseTBItems.ItemNumberSuffix = qryPremiseTBItemsCopied.ItemNumberSuffix"
    Flag =1
    LeftTable ="qryPremiseTBItems"
    RightTable ="Basic Product Construction"
    Expression ="qryPremiseTBItems.[EZ Strip] = [Basic Product Construction].[EZ Strip]"
    Flag =1
    LeftTable ="qryPremiseTBItems"
    RightTable ="Basic Product Construction"
    Expression ="qryPremiseTBItems.[TB Nominal OD] = [Basic Product Construction].[TB Nominal OD]"
    Flag =1
    LeftTable ="qryPremiseTBItems"
    RightTable ="Basic Product Construction"
    Expression ="qryPremiseTBItems.[TB Material] = [Basic Product Construction].[TB Material]"
    Flag =1
End
Begin OrderBy
    Expression ="qryPremiseTBItems.[Item No]"
    Flag =0
    Expression ="qryPremiseTBItemsCopied.ItemNumberSuffix"
    Flag =0
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    Flag =0
End
Begin Groups
    Expression ="qryPremiseTBItemsCopied.ComponentItemNumber"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[Jacket Material]"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    GroupLevel =0
    Expression ="Left(qryPremiseTBItems.[New Oracle Part #],7)"
    GroupLevel =0
    Expression ="qryPremiseTBItemsCopied.[Fiber Oracle item]"
    GroupLevel =0
    Expression ="qryPremiseTBItemsCopied.[Jacket Material]"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[Item No]"
    GroupLevel =0
    Expression ="qryPremiseTBItemsCopied.[Item No]"
    GroupLevel =0
    Expression ="Mid(qryPremiseTBItemsCopied.[Item No],12,1)"
    GroupLevel =0
    Expression ="Mid(qryPremiseTBItems![Item No],12,1)"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
    Expression ="qryPremiseTBItemsCopied.[Item No]"
    GroupLevel =0
    Expression ="Mid(qryPremiseTBItems![Item No],12,1)"
    GroupLevel =0
    Expression ="qryPremiseTBItemsCopied.ItemNumberSuffix"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    GroupLevel =0
    Expression ="Left(qryPremiseTBItems.[New Oracle Part #],7)"
    GroupLevel =0
    Expression ="Mid(qryPremiseTBItems.[Item No],7,2)"
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
    0x29e2ad462ebb81479a4b8d2e7ecde100
End
Begin
    Begin
        dbText "Name" ="NewBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x50a909f70365cc4fab1a246af4f44363
        End
    End
    Begin
        dbText "Name" ="TB"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbe5feab9c5e657428dfa9dcb8a042e57
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ddeb11a4ab1c24482b251bcf17c814d
        End
    End
    Begin
        dbText "Name" ="NewComponentItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5014cadd23fae4b972369527b063ca5
        End
    End
    Begin
        dbText "Name" ="CopiedItemNo"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f30b68b596c234e82ee66a168ac9af0
        End
    End
    Begin
        dbText "Name" ="NewItemNo"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf434f2e187da734e9792c98b0c6dada3
        End
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Fiber Oracle item]"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Item No]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-396
    Top =133
    Right =713
    Bottom =580
    Left =-1
    Top =-1
    Right =1077
    Bottom =182
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =721
        Top =16
        Right =1001
        Bottom =164
        Top =0
        Name ="qryPremiseTBItems"
        Name =""
    End
    Begin
        Left =272
        Top =25
        Right =619
        Bottom =173
        Top =0
        Name ="qryPremiseTBItemsCopied"
        Name =""
    End
    Begin
        Left =1104
        Top =8
        Right =1440
        Bottom =211
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
