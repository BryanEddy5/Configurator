Operation =3
Name ="Temp (Premise Load)"
Option =8
Where ="((([Premise Load].[Item number]) Not Like \"Shea*\") AND (([Printed Labels (Data"
    " Products)].[Completed Buffering])=False) AND (([Premise Load].[Type Order])=\"n"
    "ew\") AND (([Printed Labels (Data Products)].Shipped)=False) AND ((dbo_vScheduli"
    "ngCableRunSpeeds.CablePasses)=2))"
Begin InputTables
    Name ="Printed Labels (Data Products)"
    Name ="Premise Load"
    Name ="New Orders Spreadsheet"
    Name ="qrytotalproduced2"
    Name ="dbo_vSchedulingCableRunSpeeds"
End
Begin OutputColumns
    Alias ="LengthM"
    Name ="Length Meters"
    Expression ="IIf([Premise Load]!Units=\"Feet\",[Premise Load]!Length/3.281,[Premise Load]!Len"
        "gth)"
    Alias ="SetupHrs"
    Name ="Setup Hrs"
    Expression ="IIf([CompLength]<1,(([LengthM]*[Buffering Factor])/[PlannedLotSize])*[Set-up tim"
        "e level 2]+[Set-up time level 1],((([LengthM]*([Buffering Factor])-[CompLength])"
        "/[PlannedLotSize])*[Set-up time level 2]))"
    Alias ="Req'd Days"
    Name ="RequiredDays_Orig"
    Expression ="([B1Hrs]+[B4Hrs]+[B6Hrs]+[B2Hrs]+[B9Hrs]+[B11Hrs]+[B10Hrs]+[B18hrs])/24"
    Alias ="Adj Schedule Date"
    Name ="Adj Schedule Date"
    Expression ="[Schedule date]-[Date Adjustments]"
    Name ="Setup Buff"
    Expression ="[Printed Labels (Data Products)].[Setup Buff]"
    Name ="Setup SZ/SH"
    Expression ="[Printed Labels (Data Products)].[Setup SZ/SH]"
    Alias ="Week ending"
    Name ="Week ending"
    Expression ="[Adj Schedule date]-Weekday([Adj Schedule date],0)+6"
    Name ="Customer"
    Expression ="[Premise Load].Customer"
    Name ="Item number"
    Expression ="[Premise Load].[Item number]"
    Alias ="Expr1"
    Name ="Length"
    Expression ="[Length]/IIf([Premise Load]!Units=\"Feet\",3.28,1)"
    Name ="Mfg commit date"
    Expression ="[Premise Load].[Mfg commit date]"
    Name ="Schedule date"
    Expression ="[Premise Load].[Schedule date]"
    Name ="Reel number"
    Expression ="[Premise Load].[Co Number]"
    Alias ="Adj Mfg Date"
    Name ="Adj Mfg  Date"
    Expression ="[Mfg commit date]-[Date Adjustments]"
    Alias ="Week ending-mfg"
    Name ="Week ending-mfg"
    Expression ="[Adj Mfg Date]-Weekday([Adj Mfg Date],0)+6"
    Name ="Factor"
    Expression ="dbo_vSchedulingCableRunSpeeds.[Buffering Factor]"
    Alias ="Expr2"
    Name ="Count"
    Expression ="Mid([Item No],3,3)*1"
    Alias ="Expr3"
    Name ="Ftype"
    Expression ="Mid([Item No],6,1)"
    Name ="CO#"
    Expression ="[Premise Load].[Co Number]"
    Name ="Design"
    Expression ="dbo_vSchedulingCableRunSpeeds.ItemNo"
    Name ="Complete Buff"
    Expression ="[Printed Labels (Data Products)].[Completed Buffering]"
    Alias ="Family"
    Name ="Family"
    Expression ="Left([Item number],7)"
    Name ="Job#"
    Expression ="[Premise Load].[Job#]"
    Name ="Order Source"
    Expression ="[Premise Load].[Type Order]"
    Name ="Jacket Material"
    Expression ="dbo_vSchedulingCableRunSpeeds.JacketMaterial"
    Name ="Jacket Color"
    Expression ="dbo_vSchedulingCableRunSpeeds.Color"
    Name ="Order Qty"
    Expression ="[Printed Labels (Data Products)].[Order Qty]"
    Name ="Line#"
    Expression ="[Printed Labels (Data Products)].[Buff Line #]"
    Alias ="Expr4"
    Name ="Sequence"
    Expression ="IIf(IsNull([Run Order]),999,[Run Order])"
    Name ="Early Start"
    Expression ="[Printed Labels (Data Products)].[Early Star Date]"
    Alias ="B1hrs"
    Name ="B1 Hrs"
    Expression ="IIf([Buff Line #] In (1),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#1"
        " Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B11hrs"
    Name ="B11 Hrs"
    Expression ="IIf([Buff Line #] In (11),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#"
        "11 Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B10hrs"
    Name ="B10 Hrs"
    Expression ="IIf([Buff Line #] In (10),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#"
        "10 Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B4hrs"
    Name ="B4 Hrs"
    Expression ="IIf([Buff Line #] In (4),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#4"
        " Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B6hrs"
    Name ="B6 Hrs"
    Expression ="IIf([Buff Line #] In (6),((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#4 "
        "Line Speed]/60)/0.85+[SetupHrs],0)"
    Alias ="B2hrs"
    Name ="B2 Hrs"
    Expression ="IIf([Buff Line #] In (2),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#2"
        " Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B9hrs"
    Name ="B9 Hrs"
    Expression ="IIf([Buff Line #] In (9),(((([LengthM]*[Buffering Factor])-[CompLength])/[Buff#9"
        " Line Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="B18hrs"
    Name ="B18Hrs"
    Expression ="IIf([Buff Line #] In (18),(((([LengthM]*[Buffering Factor])-[CompLength])/[B18_L"
        "ine_Speed]/60))/0.85+[SetupHrs],0)"
    Alias ="PromDate"
    Name ="Promise Date"
    Expression ="IIf([Promise Date] Is Null,[Mfg commit date],[Promise Date])"
    Alias ="CompLength"
    Name ="CompLength"
    Expression ="IIf([SumOfOTDR Length] Is Null,0,[SumOfOTDR Length])"
    Name ="Active"
    Expression ="dbo_vSchedulingCableRunSpeeds.Active"
End
Begin Joins
    LeftTable ="Printed Labels (Data Products)"
    RightTable ="Premise Load"
    Expression ="[Printed Labels (Data Products)].[Order Qty] = [Premise Load].Length"
    Flag =1
    LeftTable ="Printed Labels (Data Products)"
    RightTable ="Premise Load"
    Expression ="[Printed Labels (Data Products)].[Reel No] = [Premise Load].[Co Number]"
    Flag =1
    LeftTable ="Premise Load"
    RightTable ="New Orders Spreadsheet"
    Expression ="[Premise Load].[Co Number] = [New Orders Spreadsheet].[Order #]"
    Flag =2
    LeftTable ="Premise Load"
    RightTable ="qrytotalproduced2"
    Expression ="[Premise Load].[Job#] = qrytotalproduced2.Job"
    Flag =2
    LeftTable ="Premise Load"
    RightTable ="dbo_vSchedulingCableRunSpeeds"
    Expression ="[Premise Load].[Item number] = dbo_vSchedulingCableRunSpeeds.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Req'd Days (#2)"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Reel number]"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Req'd Days (#4)"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Req'd Days (#9)"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week ending"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Short Date"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2cd066b4a19df946930b6892002362c0
        End
    End
    Begin
        dbText "Name" ="Adj Schedule Date"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf966dfce6200ac4a846c0ac13d3fe24a
        End
    End
    Begin
        dbText "Name" ="Premise Load.Prefix"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.Customer"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Item number]"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.Length"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Mfg commit date]"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Schedule date]"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Co number]"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Cust PO #]"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Line item]"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.[Shipped date]"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.[Cust Order Specifications].[Reel No]"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.[Cust Order Specifications].[Engr  Approval]"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.[Cust Order Specifications].[Engr  Approval - S/C]"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Premise Load.SS.Business"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="check1"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="check2"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Reel No]"
        dbInteger "ColumnWidth" ="3840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Setup Buff]"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Setup SZ/SH]"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Premise Load].Customer"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adj Mfg Date"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7eef3f8531435a4f8feec71de1fcc7ba
        End
    End
    Begin
        dbText "Name" ="Week ending-mfg"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Short Date"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x848c6b3edcebb243b26ebb618ce160e9
        End
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].Product"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Req'd Days"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x64281d1624c1924e9946be217b2778bd
        End
    End
    Begin
        dbText "Name" ="LengthM"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="0"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb29c87f1b3591a48826fb9fc239b4813
        End
    End
    Begin
        dbText "Name" ="SetupHrs"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8d41145f848e1a41bf8b6c680a76cb8c
        End
    End
    Begin
        dbText "Name" ="B1hrs"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="33"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6424bf6ad440984098b3b51f36d9f77a
        End
    End
    Begin
        dbText "Name" ="B4hrs"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="34"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9d394aee048fdc468b80e0e527d3f6fb
        End
    End
    Begin
        dbText "Name" ="B6hrs"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="35"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdbbfe96262abac48a924033e22f8cf77
        End
    End
    Begin
        dbText "Name" ="B2hrs"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="36"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7654d585946bfd4c89014e1befc4b60d
        End
    End
    Begin
        dbText "Name" ="B9hrs"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbInteger "ColumnOrder" ="37"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03151e6bafe5b64d809cc0579ccf1af2
        End
    End
    Begin
        dbText "Name" ="[Premise Load].[Item number]"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5434b1ac3bf2d9469c969834b2d6d58a
        End
    End
    Begin
        dbText "Name" ="[Premise Load].[Mfg commit date]"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Premise Load].[Schedule date]"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1012"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Buff#1/#2 Alloc Table].[Buffering Factor]"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcb758d86c35055409483d2b96aad180e
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2e321d34ef7ea74381a88a1438b23bad
        End
    End
    Begin
        dbText "Name" ="[Premise Load].[Co Number]"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Item No]"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Completed Buffering]"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Family"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7c93483ea504a94d8ae290115aba23b6
        End
    End
    Begin
        dbText "Name" ="[Premise Load].[Job#]"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Premise Load].[Type Order]"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Jacket Material]"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Jacket Color]"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Order Qty]"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Buff Line #]"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbInteger "ColumnOrder" ="31"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4f17de1da1ddf94d9a28e92e16d74f00
        End
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].[Early Star Date]"
        dbInteger "ColumnOrder" ="32"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PromDate"
        dbInteger "ColumnOrder" ="38"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e6b2cdd4807d04ba9da657a38ea40be
        End
    End
    Begin
        dbText "Name" ="CompLength"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8aaaf5a6d9291843ad842361714f83ee
        End
    End
    Begin
        dbText "Name" ="B11hrs"
        dbInteger "ColumnWidth" ="675"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc3cc617e70924c4fb7f8614d4099f2db
        End
    End
    Begin
        dbText "Name" ="B10hrs"
        dbInteger "ColumnWidth" ="675"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x50a29177bf36fd4cb3cc0c7990d00e09
        End
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.[Jacket Color]"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Color"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.[Jacket Material]"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.[Item No]"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.ItemNo"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Printed Labels (Data Products)].Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.CenterMember"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.[Buffering Factor]"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.ItemNo"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.JacketMaterial"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.Color"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vSchedulingCableRunSpeeds.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B18hrs"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8e8d73982e956143b72f6882aaa5abdb
        End
    End
End
Begin
    State =0
    Left =-111
    Top =44
    Right =1341
    Bottom =854
    Left =-1
    Top =-1
    Right =1420
    Bottom =276
    Left =0
    Top =12
    ColumnsShown =651
    Begin
        Left =59
        Top =24
        Right =155
        Bottom =168
        Top =0
        Name ="Printed Labels (Data Products)"
        Name =""
    End
    Begin
        Left =219
        Top =12
        Right =365
        Bottom =156
        Top =0
        Name ="Premise Load"
        Name =""
    End
    Begin
        Left =563
        Top =215
        Right =744
        Bottom =360
        Top =0
        Name ="New Orders Spreadsheet"
        Name =""
    End
    Begin
        Left =258
        Top =233
        Right =354
        Bottom =317
        Top =0
        Name ="qrytotalproduced2"
        Name =""
    End
    Begin
        Left =645
        Top =11
        Right =1102
        Bottom =208
        Top =0
        Name ="dbo_vSchedulingCableRunSpeeds"
        Name =""
    End
End
