Operation =4
Option =0
Where ="(((tblColors_Available.Color_ID)=[Forms]![frmCreateItem]![SubColor]))"
Begin InputTables
    Name ="tblNewPrductConstruction"
    Name ="tblCableConstructionReferences"
    Name ="tblColors_Available"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[New Oracle Part #]"
    Expression ="NextItemNo(\"EE\",True) & \"-00\""
    Name ="tblNewPrductConstruction.[Item No]"
    Expression ="Left([tblCableConstructionReferences].[DesignCode],5) & [Forms]![frmCreateItem]!"
        "[FiberType] & Mid([tblCableConstructionReferences].[DesignCode],6,2) & \"U\" & ["
        "Forms]![frmCreateItem]![SubColor] & [Forms]![frmCreateItem]![TBColor] & [Forms]!"
        "[frmCreateItem]![TBType] & [DesignExtension] & Mid([Item No],13,7)"
    Name ="tblNewPrductConstruction.Active"
    Expression ="-1"
    Name ="tblNewPrductConstruction.Reason"
    Expression ="\"IP \" & [Forms]![frmSwitchBoard]![Initials]"
    Name ="tblNewPrductConstruction.[Revision Date]"
    Expression ="Date()"
    Name ="tblNewPrductConstruction.[Revision Letter]"
    Expression ="\"1\""
    Name ="tblNewPrductConstruction.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & Forms!frmSwitchboard!Initials"
    Name ="tblNewPrductConstruction.[Fiber Type]"
    Expression ="Forms!frmFiberSpec!FiberType"
    Name ="tblNewPrductConstruction.[Fiber Oracle item]"
    Expression ="Forms!frmFiberSpec!Fiber"
    Name ="tblNewPrductConstruction.[1st Req Freq]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewPrductConstruction.[1st Max Atten]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewPrductConstruction.[1st Min BandW]"
    Expression ="Forms!frmFiberSpec![1stMinBandW]"
    Name ="tblNewPrductConstruction.[2nd Req Freq]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewPrductConstruction.[2nd Max Atten]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewPrductConstruction.[2nd Min BandW]"
    Expression ="Forms!frmFiberSpec![2ndMinBandW]"
    Name ="tblNewPrductConstruction.[SM 1300 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.[SM 1550 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.OracleStatus"
    Expression ="\"Acitve\""
    Name ="tblNewPrductConstruction.Customer"
    Expression ="\"AFL STANDARD\""
    Name ="tblNewPrductConstruction.FiberType2"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.FiberType3"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.[Jacket Color]"
    Expression ="[tblColors_Available].[Color]"
End
Begin Joins
    LeftTable ="tblNewPrductConstruction"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblNewPrductConstruction.Base = tblCableConstructionReferences.Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xaefb4ae8ccdbbf45a921262512875d06
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[SM 1550 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Revision History]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Item No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Fiber Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[1st Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[1st Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[1st Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[2nd Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[2nd Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[2nd Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[SM 1300 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblColors_Available.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblColors_Available.Color_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1007"
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
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =90
        Top =64
        Right =457
        Bottom =374
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
    Begin
        Left =532
        Top =18
        Right =984
        Bottom =300
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1034
        Top =20
        Right =1316
        Bottom =295
        Top =0
        Name ="tblColors_Available"
        Name =""
    End
End
