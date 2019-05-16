Operation =1
Option =0
Where ="(((tblBOMOpResourcesInterface.ResourceCode)=\"setup\"))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
    Name ="tblBOMOpResourcesInterface"
End
Begin OutputColumns
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="tblBOMOpResourcesInterface.ResourceCode"
    Expression ="tblBOMOpSequencesInterface.DepartmentCode"
    Alias ="StopLoad"
    Expression ="0"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterface"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterface.OrganizationCode = tblBOMOpResourcesInterface.Organiz"
        "ationCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterface"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterface.AssemblyItemNumber = tblBOMOpResourcesInterface.Assem"
        "blyItemNumber"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterface"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterface.OperationSeqNum = tblBOMOpResourcesInterface.Operatio"
        "nSeqNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xde7c02f03ab9a8488c10087b02ee1474
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3045671d7953fe409b9de30b5a61e8fa
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
    Bottom =291
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =625
        Top =31
        Right =851
        Bottom =220
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
    Begin
        Left =140
        Top =31
        Right =451
        Bottom =220
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
End
