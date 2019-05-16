dbMemo "SQL" ="PARAMETERS [Forms]![frmMainMenu]![ItemSetupOrg] Text ( 255 );\015\012SELECT tblP"
    "roductCategorySegment1.Segment1Desc, tblProductCategorySegment2.*, tblProductCat"
    "egorySegment1!Segment1ID & \".\" & [Segment2ID] & \".0.0.0\" AS ProductType, [Se"
    "gment1Desc] & \".\" & [Segment2Desc] AS ProductTypeView, tblProductCategorySegme"
    "nt1.Active\015\012FROM tblProductCategorySegment1 INNER JOIN tblProductCategoryS"
    "egment2 ON tblProductCategorySegment1.Segment1ID = tblProductCategorySegment2.Se"
    "gment1ID\015\012WHERE (((tblProductCategorySegment1.Segment1Desc)=Forms!frmMainM"
    "enu!ItemSetupOrg) And ((tblProductCategorySegment1.Active)=True) And ((tblProduc"
    "tCategorySegment2.Active)=True));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x8e22d744e0571e42aadc70a10d301021
End
Begin
    Begin
        dbText "Name" ="ProductType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1b363d1f20c664409bb1962b8792eb1f
        End
    End
    Begin
        dbText "Name" ="ProductTypeView"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40bfb88f88f9fd4cb3a8ee56e2529d12
        End
    End
    Begin
        dbText "Name" ="tblProductCategorySegment1.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment1.Segment1Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.Segment1ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.Segment2ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.Segment2Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.NewSwg2ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProductCategorySegment2.NewSeg1ID"
        dbLong "AggregateType" ="-1"
    End
End
