Operation =1
Option =0
Having ="((([Basic Product Construction].[New Oracle Part #])=[Forms]![frmItemInterface]!"
    "[frmPremiseItemDetails].[Form]![OracleItemNumber]))"
Begin InputTables
    Name ="qryPremiseTBItems"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    Expression ="qryPremiseTBItems.[Jacket Material]"
    Alias ="NewComponentItem"
    Expression ="Left(qryPremiseTBItems.[New Oracle Part #],7)"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
Begin Joins
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
    Expression ="qryPremiseTBItems.[Jacket Material] = [Basic Product Construction].[TB Material]"
    Flag =1
End
Begin OrderBy
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    Flag =0
End
Begin Groups
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[Jacket Material]"
    GroupLevel =0
    Expression ="Left(qryPremiseTBItems.[New Oracle Part #],7)"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc6299e0c98effb49b796582b658e630b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewComponentItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8bc08897bee77a4c8350f92a1827fccc
        End
    End
End
Begin
    State =0
    Left =-218
    Top =130
    Right =891
    Bottom =577
    Left =-1
    Top =-1
    Right =1077
    Bottom =233
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =728
        Top =19
        Right =1008
        Bottom =167
        Top =0
        Name ="qryPremiseTBItems"
        Name =""
    End
    Begin
        Left =1161
        Top =30
        Right =1497
        Bottom =178
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
