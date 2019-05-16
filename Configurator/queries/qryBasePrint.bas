dbMemo "SQL" ="SELECT [Print Line 1] AS PrintLine1, Trim([Customer]) AS Cust, [Print Line 2] AS"
    " PrintLine2, [Print Line 3] AS PrintLine3, LEFT([Item No],2) AS CabFam, [New Ora"
    "cle Part #] AS Oracle, [Print Line 4] AS PrintLine4, [Print Item No], [Print Typ"
    "e (base)], [Print Reel No], [Item No]\015\012FROM [Basic Product Construction];\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x439ca51ac986b642a21e90d421fd3093
End
dbMemo "Filter" ="([qryBasePrint].[PrintLine3] Like \"*made*\")"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "OrderBy" ="[qryBasePrint].[PrintLine3], [qryBasePrint].[Oracle]"
Begin
    Begin
        dbText "Name" ="PrintLine1"
        dbInteger "ColumnWidth" ="5460"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x941ce0d4f6a9c445afc3e1f78d871f13
        End
    End
    Begin
        dbText "Name" ="PrintLine2"
        dbInteger "ColumnWidth" ="4785"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3982595eba6af64ea0bcdb68681af076
        End
    End
    Begin
        dbText "Name" ="PrintLine3"
        dbInteger "ColumnWidth" ="6600"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe9088e9988d00a409b03cd07886fc7d2
        End
    End
    Begin
        dbText "Name" ="CabFam"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdaa6e18a2256364686bf5d145ab90c28
        End
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Cust"
        dbInteger "ColumnWidth" ="2940"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x630f8fa143d7224996de6e41fd081efe
        End
    End
    Begin
        dbText "Name" ="Oracle"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x39e07181b38ae643a6e71b58f279fec3
        End
    End
    Begin
        dbText "Name" ="PrintLine4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8fbf2f5fd0f09b4f9a753f843b84f726
        End
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[Print Type (base)]"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Print Item No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa812778291a56a4b8e6fcbea53405366
        End
    End
    Begin
        dbText "Name" ="Print Type (base)"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdc2b9036b9577b47850bb2cf0de7a9e2
        End
    End
    Begin
        dbText "Name" ="Print Reel No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2146d374614af645b06bd590f448f163
        End
    End
    Begin
        dbText "Name" ="Item No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa4cb9dee5c72fa4f9ebaf5f2ed0e4526
        End
    End
    Begin
        dbText "Name" ="[Print Reel No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Print Item No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
End
