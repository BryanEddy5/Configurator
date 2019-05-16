Operation =1
Option =0
Where ="((([tblOracleBOMOpSequencesInterface].[AssemblyItemNumber])=[Forms]![frmMainMenu"
    "]![ItemNumberSelect]))"
Begin InputTables
    Name ="tblOracleBOMOpSequencesInterface"
End
Begin OutputColumns
    Alias ="ExistingRoutingCheck"
    Expression ="Count(tblOracleBOMOpSequencesInterface.AssemblyItemNumber)"
    Alias ="MissingElement"
    Expression ="IIf(Count([OrganizationCode])=0,-1,0)"
    Alias ="MissingElementType"
    Expression ="IIf(Count([OrganizationCode])=0,\"Routing\",\"\")"
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
    0x0944ab5fe5ac49448196af37eca908be
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingRoutingCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf65e19bf54e22243933d80a7754b0dd7
        End
    End
    Begin
        dbText "Name" ="MissingElement"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6f11d504dde8054cb225ef511ba08170
        End
    End
    Begin
        dbText "Name" ="MissingElementType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c0422d05ce3804a8427338c053b6157
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
        Right =285
        Bottom =120
        Top =0
        Name ="tblOracleBOMOpSequencesInterface"
        Name =""
    End
End
