Operation =1
Option =0
Where ="((([tblOracleItemCategoryInterface].[ItemNumber])=[Forms]![frmMainMenu]![ItemNum"
    "berSelect]))"
Begin InputTables
    Name ="tblOracleItemCategoryInterface"
End
Begin OutputColumns
    Alias ="ExistingCategoryCheck"
    Expression ="Count(tblOracleItemCategoryInterface.ItemNumber)"
    Alias ="MissingElement"
    Expression ="IIf(Count([OrganizationCode])=0,-1,0)"
    Alias ="MissingElementType"
    Expression ="IIf(Count([OrganizationCode])=0,\"Category Information\",\"\")"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x98640637869b724fb965f04a3248e2d4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingCategoryCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x27c8a436cb63a047889e6d84d8659060
        End
    End
    Begin
        dbText "Name" ="MissingElementType"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5aacc500dd228143beb8e167031198a8
        End
    End
    Begin
        dbText "Name" ="MissingElement"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7fa7022b61e17b4f8ecbb670eee66660
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =366
        Top =6
        Right =617
        Bottom =120
        Top =0
        Name ="tblOracleItemCategoryInterface"
        Name =""
    End
End
