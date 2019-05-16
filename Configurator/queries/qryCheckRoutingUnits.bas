Operation =1
Option =0
Where ="(((tblBOMOpResourcesInterface.AssignedUnits) Is Null))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMOpResourcesInterface"
    Name ="tblOracleResourceCodes"
End
Begin OutputColumns
    Alias ="ComponentItemNumber"
    Expression ="\"\""
    Alias ="ItemSeqNumber"
    Expression ="\"\""
    Alias ="CheckCode"
    Expression ="\"Routing - Missing Assigned Units, Resource: \" & tblBOMOpResourcesInterface!Re"
        "sourceCode"
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="tblBOMOpResourcesInterface.UsageRateOrAmount"
    Expression ="tblBOMOpResourcesInterface.ResourceCode"
    Expression ="tblOracleResourceCodes.AllowZeroUsage"
    Expression ="tblBOMOpResourcesInterface.AssignedUnits"
    Alias ="StopLoad"
    Expression ="-1"
End
Begin Joins
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblOracleResourceCodes"
    Expression ="tblBOMOpResourcesInterface.ResourceCode = tblOracleResourceCodes.ResourceCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode = tblBOMOpResourcesInterface.Org"
        "anizationCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber = tblBOMOpResourcesInterface.A"
        "ssemblyItemNumber"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum = tblBOMOpResourcesInterface.Oper"
        "ationSeqNumber"
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
    0xb9e0b5abd238be4db6a0dc95bc11f268
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x27a85c079e8c08489162a88d05f8a164
        End
    End
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb117f564c6b63b4b947f33fa50a4ba7e
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x32477367a305394dbbd864ada1b1b098
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41994ebe381990468b526d080c0148a7
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1418
    Bottom =443
    Left =-1
    Top =-1
    Right =1368
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =34
        Top =10
        Right =318
        Bottom =184
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =509
        Top =6
        Right =820
        Bottom =195
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
    Begin
        Left =933
        Top =14
        Right =1094
        Bottom =158
        Top =0
        Name ="tblOracleResourceCodes"
        Name =""
    End
End
