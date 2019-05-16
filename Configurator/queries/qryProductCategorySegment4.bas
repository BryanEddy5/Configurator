dbMemo "SQL" ="SELECT qryProductCategorySegment3.Segment1Desc, qryProductCategorySegment3.Segme"
    "nt2Desc, qryProductCategorySegment3.Segment3Desc, tblProductCategorySegment4.*, "
    "tblProductCategorySegment4!Segment1ID & \".\" & tblProductCategorySegment4!Segme"
    "nt2ID & \".\" & tblProductCategorySegment4!Segment3ID & \".\" & [Segment4ID] & \""
    ".0\" AS ProductType, [Segment1Desc] & \".\" & [Segment2Desc] & \".\" & [Segment3"
    "Desc] & \".\" & [Segment4Desc] AS ProductTypeView, qryProductCategorySegment3.tb"
    "lProductCategorySegment1.Active, qryProductCategorySegment3.tblProductCategorySe"
    "gment2.Active, qryProductCategorySegment3.tblProductCategorySegment3.Active\015\012"
    "FROM tblProductCategorySegment4 INNER JOIN qryProductCategorySegment3 ON (tblPro"
    "ductCategorySegment4.Segment1ID=qryProductCategorySegment3.Segment1ID) AND (tblP"
    "roductCategorySegment4.Segment2ID=qryProductCategorySegment3.Segment2ID) AND (tb"
    "lProductCategorySegment4.Segment3ID=qryProductCategorySegment3.Segment3ID)\015\012"
    "WHERE (((qryProductCategorySegment3.tblProductCategorySegment1.Active)=-1) AND ("
    "(qryProductCategorySegment3.tblProductCategorySegment2.Active)=-1) AND ((qryProd"
    "uctCategorySegment3.tblProductCategorySegment3.Active)=-1) AND ((tblProductCateg"
    "orySegment4.Active)=-1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2164cd6b88a97d44904e47fe768bbf9e
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
            0xce5588013ffb2e448a55dceec330eb03
        End
    End
    Begin
        dbText "Name" ="ProductTypeView"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b5fd6a2ea48544ebff67bb174a6e9c6
        End
    End
End
