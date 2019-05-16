dbMemo "SQL" ="SELECT qryProductCategorySegment4.Segment1Desc, qryProductCategorySegment4.Segme"
    "nt2Desc, qryProductCategorySegment4.Segment3Desc, qryProductCategorySegment4.Seg"
    "ment4Desc, tblProductCategorySegment5.*, tblProductCategorySegment5!Segment1ID &"
    " \".\" & tblProductCategorySegment5!Segment2ID & \".\" & tblProductCategorySegme"
    "nt5!Segment3ID & \".\" & tblProductCategorySegment5!Segment4ID & \".\" & [Segmen"
    "t5ID] AS ProductType, [Segment1Desc] & \".\" & [Segment2Desc] & \".\" & [Segment"
    "3Desc] & \".\" & [Segment4Desc] & \".\" & [Segment5Desc] AS ProductTypeView, qry"
    "ProductCategorySegment4.tblProductCategorySegment1.Active, qryProductCategorySeg"
    "ment4.tblProductCategorySegment2.Active, qryProductCategorySegment4.tblProductCa"
    "tegorySegment3.Active, qryProductCategorySegment4.tblProductCategorySegment4.Act"
    "ive\015\012FROM tblProductCategorySegment5 INNER JOIN qryProductCategorySegment4"
    " ON (tblProductCategorySegment5.Segment1ID=qryProductCategorySegment4.Segment1ID"
    ") AND (tblProductCategorySegment5.Segment2ID=qryProductCategorySegment4.Segment2"
    "ID) AND (tblProductCategorySegment5.Segment3ID=qryProductCategorySegment4.Segmen"
    "t3ID) AND (tblProductCategorySegment5.Segment4ID=qryProductCategorySegment4.Segm"
    "ent4ID);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x819339219524754787f8f2538d199160
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ProductType"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x20200e84eaef70469674177c501d6ba4
        End
    End
    Begin
        dbText "Name" ="ProductTypeView"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bd7cd637e7e7544bbdde84cc1b86ade
        End
    End
End
