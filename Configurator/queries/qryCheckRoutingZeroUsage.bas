Operation =1
Option =0
Where ="(((tblBOMOpResourcesInterface.UsageRateOrAmount)=0) AND ((tblOracleResourceCodes"
    ".AllowZeroUsage)=False))"
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
    Expression ="\"Routing - Zero Usage, Resource: \" & tblBOMOpResourcesInterface!ResourceCode"
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="tblBOMOpResourcesInterface.UsageRateOrAmount"
    Expression ="tblBOMOpResourcesInterface.ResourceCode"
    Expression ="tblOracleResourceCodes.AllowZeroUsage"
    Alias ="StopLoad"
    Expression ="0"
End
Begin Joins
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblOracleResourceCodes"
    Expression ="tblBOMOpResourcesInterface.ResourceCode = tblOracleResourceCodes.ResourceCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblOracleResourceCodes"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode = tblOracleResourceCodes.Departmen"
        "tCode"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum = tblBOMOpResourcesInterface.Oper"
        "ationSeqNumber"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber = tblBOMOpResourcesInterface.A"
        "ssemblyItemNumber"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode = tblBOMOpResourcesInterface.Org"
        "anizationCode"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1fb862e793e82749a2e0a302a64787d7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="3660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5c515b2fab676e4186eccc65c22c7e8c
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
            0xbb66dc9e048beb449653d6d471356502
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb5a17adc861ca144ad3a383c08506847
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa569ce2f21ce5498ab4f8ff62580400
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
        Left =56
        Top =5
        Right =301
        Bottom =194
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =537
        Top =5
        Right =848
        Bottom =194
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
    Begin
        Left =997
        Top =9
        Right =1158
        Bottom =153
        Top =0
        Name ="tblOracleResourceCodes"
        Name =""
    End
End
