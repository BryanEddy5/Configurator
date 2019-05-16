Operation =1
Option =0
Where ="((([tblOracleBOMInventoryCompsInterface].[AssemblyItemNumber])=[Forms]![frmMainM"
    "enu]![ItemNumberSelect]))"
Begin InputTables
    Name ="tblOracleBOMInventoryCompsInterface"
End
Begin OutputColumns
    Alias ="ExistingBOMCheck"
    Expression ="Count(tblOracleBOMInventoryCompsInterface.AssemblyItemNumber)"
    Alias ="MissingElement"
    Expression ="IIf(Count([OrganizationCode])=0,-1,0)"
    Alias ="MissingElementType"
    Expression ="IIf(Count([OrganizationCode])=0,\"Bill of Materials\",\"\")"
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
    0x1f78e408abf8fd48b7744fd6f33af0c4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingBOMCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9950e4befc32ac48b2c1dbea2b350ef2
        End
    End
    Begin
        dbText "Name" ="MissingElement"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x51f43aeeba813443967e758516e21646
        End
    End
    Begin
        dbText "Name" ="MissingElementType"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b2b311823cdcd40ad6b711d28f92f9e
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
        Left =38
        Top =6
        Right =328
        Bottom =120
        Top =0
        Name ="tblOracleBOMInventoryCompsInterface"
        Name =""
    End
End
