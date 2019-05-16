dbMemo "SQL" ="SELECT tblTestBOMInventoryCompsInterface.AssemblyItemNumber, tblNetworkBOMInvent"
    "oryCompsInterface.AssemblyItemNumber, tblTestBOMInventoryCompsInterface.Componen"
    "tItemNumber, tblNetworkBOMInventoryCompsInterface.ComponentItemNumber, tblTestBO"
    "MInventoryCompsInterface.OperationSeqNumber, tblNetworkBOMInventoryCompsInterfac"
    "e.OperationSeqNumber, tblTestBOMInventoryCompsInterface.ComponentQuantity, tblNe"
    "tworkBOMInventoryCompsInterface.ComponentQuantity, tblTestBOMInventoryCompsInter"
    "face.ItemSeqNumber, tblNetworkBOMInventoryCompsInterface.ItemSeqNumber, IIf(tblT"
    "estBOMInventoryCompsInterface!ComponentItemNumber<>tblNetworkBOMInventoryCompsIn"
    "terface!ComponentItemNumber,-1,0) AS Expr1, IIf(tblTestBOMInventoryCompsInterfac"
    "e!OperationSeqNumber<>tblNetworkBOMInventoryCompsInterface!OperationSeqNumber,-1"
    ",0) AS Expr2, IIf(tblTestBOMInventoryCompsInterface!ComponentQuantity<>tblNetwor"
    "kBOMInventoryCompsInterface!ComponentQuantity,-1,0) AS Expr3, IIf(tblTestBOMInve"
    "ntoryCompsInterface!ItemSeqNumber<>tblNetworkBOMInventoryCompsInterface!ItemSeqN"
    "umber,-1,0) AS Expr4\015\012FROM tblNetworkBOMInventoryCompsInterface RIGHT JOIN"
    " tblTestBOMInventoryCompsInterface ON (tblNetworkBOMInventoryCompsInterface.Orga"
    "nizationCode=tblTestBOMInventoryCompsInterface.OrganizationCode) AND (tblNetwork"
    "BOMInventoryCompsInterface.AssemblyItemNumber=tblTestBOMInventoryCompsInterface."
    "AssemblyItemNumber) AND (tblNetworkBOMInventoryCompsInterface.ComponentItemNumbe"
    "r=tblTestBOMInventoryCompsInterface.ComponentItemNumber) AND (tblNetworkBOMInven"
    "toryCompsInterface.OperationSeqNumber=tblTestBOMInventoryCompsInterface.Operatio"
    "nSeqNumber);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="(((tblTestBOMInventoryCompsInterface.AssemblyItemNumber=\"NEW ITEM 1W\"))) AND ("
    "(tblTestBOMInventoryCompsInterface.OperationSeqNumber=190))"
dbMemo "OrderBy" ="tblTestBOMInventoryCompsInterface.ItemSeqNumber"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0a95a2b91764834c8c2bba323478fdba
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.AssemblyItemNumber"
        dbInteger "ColumnWidth" ="5640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMInventoryCompsInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="5250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMInventoryCompsInterface.ComponentQuantity"
        dbInteger "ColumnWidth" ="5115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ComponentQuantity"
        dbInteger "ColumnWidth" ="5445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMInventoryCompsInterface.ItemSeqNumber"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMInventoryCompsInterface.ItemSeqNumber"
        dbInteger "ColumnWidth" ="5115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9bec6bcdae394143ad115a9fece43662
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xab6971e6293c2245ae16983e08347941
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a6c192386a00a44acaf528aab590fc9
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb6cc54be1e986c489988fee9aa7b438e
        End
    End
End
