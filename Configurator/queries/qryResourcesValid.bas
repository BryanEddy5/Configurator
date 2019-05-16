Operation =1
Option =0
Begin InputTables
    Name ="tblBOMOpResourcesInterface"
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblOracleResourceCodes"
End
Begin OutputColumns
    Expression ="tblBOMOpResourcesInterface.*"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode"
End
Begin Joins
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber = tblBOMOpSequencesInterfaceTemp.O"
        "perationSeqNum"
    Flag =1
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMOpResourcesInterface.AssemblyItemNumber = tblBOMOpSequencesInterfaceTemp.A"
        "ssemblyItemNumber"
    Flag =1
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMOpResourcesInterface.OrganizationCode = tblBOMOpSequencesInterfaceTemp.Org"
        "anizationCode"
    Flag =1
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblOracleResourceCodes"
    Expression ="tblBOMOpResourcesInterface.ResourceCode = tblOracleResourceCodes.ResourceCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblOracleResourceCodes"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode = tblOracleResourceCodes.Departmen"
        "tCode"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbd81659f32c9aa48bfcb32cc39da079d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.ResourceSeqNumber"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =389
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =133
        Top =17
        Right =381
        Bottom =372
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
    Begin
        Left =773
        Top =27
        Right =1067
        Bottom =322
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =501
        Top =132
        Right =707
        Bottom =361
        Top =0
        Name ="tblOracleResourceCodes"
        Name =""
    End
End
