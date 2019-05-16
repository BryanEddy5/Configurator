Operation =1
Option =2
Where ="(((tblSalesForceItemData_1.[Task Subject])=\"Item Set-up Request Review\") AND ("
    "(tblSalesForceItemData_1.Status)<>\"Cancelled\") AND ((tblSalesForceItemData.Sta"
    "tus)<>\"Cancelled\") AND ((tblSalesForceItemData.[Task Subject])=\"Cable Item Se"
    "t-up\"))"
Begin InputTables
    Name ="tblSalesForceItemData"
    Name ="tblSalesForceItemData"
    Alias ="tblSalesForceItemData_1"
End
Begin OutputColumns
    Expression ="tblSalesForceItemData.[Related To]"
    Alias ="TaskOpened"
    Expression ="tblSalesForceItemData.[Task Closed Date]"
    Expression ="tblSalesForceItemData_1.[Task Subject]"
    Alias ="TaskClosed"
    Expression ="tblSalesForceItemData_1.[Task Closed Date]"
    Expression ="tblSalesForceItemData_1.Assigned"
    Alias ="DaysToCloseTask"
    Expression ="CalculateDaysOpen([TaskOpened],[TaskClosed],False)"
    Alias ="Date"
    Expression ="Format([tblSalesForceItemData_1].[Date],\"yyyy/mm\")"
    Expression ="tblSalesForceItemData_1.[Task Waiting Reason]"
    Alias ="CompletedIn48Hours"
    Expression ="IIf([DaysToCloseTask]>=2,\"No\",\"Yes\")"
End
Begin Joins
    LeftTable ="tblSalesForceItemData"
    RightTable ="tblSalesForceItemData_1"
    Expression ="tblSalesForceItemData.[Related To] = tblSalesForceItemData_1.[Related To]"
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
    0xff9b93926e362943901ac29390510f39
End
Begin
    Begin
        dbText "Name" ="tblSalesForceItemData.Related To"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForceItemData.[Related To]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForceItemData_1.[Task Subject]"
        dbInteger "ColumnWidth" ="3930"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForceItemData_1.Assigned"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblSalesForceItemData_1.[Task Waiting Reason]"
        dbInteger "ColumnWidth" ="4920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="DaysToCloseTask"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6506a2bee71c3441ac662e2c0d09aad2
        End
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="TaskOpened"
        dbInteger "ColumnWidth" ="2280"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6db31cea3ce8ed40a143ddf8f98b1dff
        End
    End
    Begin
        dbText "Name" ="TaskClosed"
        dbInteger "ColumnWidth" ="3375"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xce393dbbe37c864682d9804d7d87abb3
        End
    End
    Begin
        dbText "Name" ="CompletedIn48Hours"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5591e4402fcfc042b2657b8321979680
        End
    End
End
Begin
    State =0
    Left =36
    Top =145
    Right =1145
    Bottom =774
    Left =-1
    Top =-1
    Right =1077
    Bottom =307
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =241
        Bottom =222
        Top =0
        Name ="tblSalesForceItemData"
        Name =""
    End
    Begin
        Left =342
        Top =4
        Right =577
        Bottom =314
        Top =0
        Name ="tblSalesForceItemData_1"
        Name =""
    End
End
