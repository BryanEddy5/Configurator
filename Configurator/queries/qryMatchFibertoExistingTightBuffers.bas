Operation =1
Option =2
Where ="(((qryFiberOracle.Oracle) Not Like \"*TEST*\") And ((qryFiberOracle.[Item No]) L"
    "ike Left(Forms!frmMainMenu!ItemNo,8) & \"??\" & Mid(Forms!frmMainMenu!ItemNo,11,"
    "2) & \"*\") And ((qryFiberOracle.Active)<>0))"
Begin InputTables
    Name ="qryFiberOracle"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.Fiber"
    Expression ="qryFiberOracle.Color"
    Expression ="qryFiberOracle.Jacket"
    Expression ="qryFiberOracle.Base"
    Expression ="qryFiberOracle.ColorChip"
    Expression ="qryFiberOracle.Customer"
    Expression ="qryFiberOracle.PrintLine4"
    Alias ="Active1"
    Expression ="IIf(qryFiberOracle!Active=0,\"No\",\"Yes\")"
    Expression ="qryFiberOracle.Reason"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Fiber2"
    Expression ="qryFiberOracle.Location"
End
Begin OrderBy
    Expression ="qryFiberOracle.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x584d1bb2761a1f43afa05b11fe30f175
End
dbMemo "OrderBy" ="qryMatchFibertoExistingTightBuffers.Oracle"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Active1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x68d60425cac8bd46bc587cc0d5103aa4
        End
    End
End
Begin
    State =0
    Left =86
    Top =47
    Right =1378
    Bottom =617
    Left =-1
    Top =-1
    Right =1260
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =501
        Top =19
        Right =712
        Bottom =182
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
End
