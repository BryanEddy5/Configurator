Operation =1
Option =0
Where ="(((qryResourcesValid.ResourceCode) Is Null) AND ((tblBOMOpResourcesInterface.Res"
    "ourceCode)<>\"0\"))"
Begin InputTables
    Name ="tblBOMOpResourcesInterface"
    Name ="qryResourcesValid"
End
Begin OutputColumns
    Expression ="tblBOMOpResourcesInterface.*"
    Alias ="ComponentItemNumber"
    Expression ="\"\""
    Alias ="ItemSeqNumber"
    Expression ="\"\""
    Alias ="CheckCode"
    Expression ="\"Routing - Invalid Resource \" & tblBOMOpResourcesInterface!ResourceCode"
    Alias ="StopLoad"
    Expression ="-1"
End
Begin Joins
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="qryResourcesValid"
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber = qryResourcesValid.OperationSeqNu"
        "mber"
    Flag =2
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="qryResourcesValid"
    Expression ="tblBOMOpResourcesInterface.AssemblyItemNumber = qryResourcesValid.AssemblyItemNu"
        "mber"
    Flag =2
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="qryResourcesValid"
    Expression ="tblBOMOpResourcesInterface.OrganizationCode = qryResourcesValid.OrganizationCode"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x194af50e69b1d84391471ffcd4370a5c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.ResourceCode"
        dbInteger "ColumnWidth" ="4095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2933caeeaedec74e9a108cefbf643e99
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
            0x193a4b63ddf7c8408871873ceea85d4b
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf521bbb56c1abd46abf6be77a2921b7e
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4b67b793a15ac74e9a66c9b71b8e48ec
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
    Bottom =325
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =420
        Bottom =300
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
    Begin
        Left =537
        Top =9
        Right =728
        Bottom =273
        Top =0
        Name ="qryResourcesValid"
        Name =""
    End
End
