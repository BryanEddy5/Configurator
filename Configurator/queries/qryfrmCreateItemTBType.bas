Operation =1
Option =0
Having ="(((qryfrmCreateItemTBType_NoTB.[TB Material]) Not Like \"frc*\") AND ((qryfrmCre"
    "ateItemTBType_NoTB.Active)=Yes))"
Begin InputTables
    Name ="qryfrmCreateItemTBType_NoTB"
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Alias ="Family"
    Expression ="Left([qryfrmCreateItemTBType_NoTB].[Item No],2)"
    Alias ="TBType"
    Expression ="Mid([qryfrmCreateItemTBType_NoTB].[Item No],12,1)"
    Expression ="qryfrmCreateItemTBType_NoTB.JacketMaterial"
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Material]"
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Chips Type]"
    Alias ="CountOfTB Material"
    Expression ="Count(qryfrmCreateItemTBType_NoTB.[TB Material])"
    Expression ="qryfrmCreateItemTBType_NoTB.[EZ Strip]"
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Nominal OD]"
End
Begin Joins
    LeftTable ="qryfrmCreateItemTBType_NoTB"
    RightTable ="qryItemNoJoin"
    Expression ="qryfrmCreateItemTBType_NoTB.[New Oracle Part #] = qryItemNoJoin.Oracle"
    Flag =1
End
Begin OrderBy
    Expression ="Left([qryfrmCreateItemTBType_NoTB].[Item No],2)"
    Flag =0
    Expression ="Mid([qryfrmCreateItemTBType_NoTB].[Item No],12,1)"
    Flag =0
    Expression ="qryfrmCreateItemTBType_NoTB.JacketMaterial"
    Flag =0
    Expression ="Count(qryfrmCreateItemTBType_NoTB.[TB Material])"
    Flag =1
End
Begin Groups
    Expression ="Left([qryfrmCreateItemTBType_NoTB].[Item No],2)"
    GroupLevel =0
    Expression ="Mid([qryfrmCreateItemTBType_NoTB].[Item No],12,1)"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.JacketMaterial"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Material]"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Chips Type]"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.[EZ Strip]"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.[TB Nominal OD]"
    GroupLevel =0
    Expression ="qryfrmCreateItemTBType_NoTB.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe22bed252451894dac797fd2516519bb
End
dbInteger "RowHeight" ="270"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="((([qryfrmCreateItemTBType].[TBType]=\"1\") AND ([qryfrmCreateItemTBType].[Famil"
    "y]=\"ER\"))) AND ([qryfrmCreateItemTBType].[JacketMaterial]=\"PVC00117\")"
dbMemo "OrderBy" ="[qryfrmCreateItemTBType].[TBType]"
Begin
    Begin
        dbText "Name" ="TBType"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1dc2ee6ef1da744cb69340a563471154
        End
    End
    Begin
        dbText "Name" ="CountOfTB Material"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xae46a603be9f3f4bae2ea9d3a175b70d
        End
    End
    Begin
        dbText "Name" ="Family"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb99f72dceb38084197c3012128eb009c
        End
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_NoTB.[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_NoTB.[TB Material]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_NoTB.[TB Chips Type]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_NoTB.[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryfrmCreateItemTBType_NoTB.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =145
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =286
        Top =24
        Right =532
        Bottom =167
        Top =0
        Name ="qryfrmCreateItemTBType_NoTB"
        Name =""
    End
    Begin
        Left =643
        Top =4
        Right =957
        Bottom =167
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End
