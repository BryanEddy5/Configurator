dbMemo "SQL" ="SELECT qryProductCategorySegment2.Segment1Desc, qryProductCategorySegment2.Segme"
    "nt2Desc, tblProductCategorySegment3.*, tblProductCategorySegment3!Segment1ID & \""
    ".\" & tblProductCategorySegment3!Segment2ID & \".\" & [Segment3ID] & \".0.0\" AS"
    " ProductType, [Segment1Desc] & \".\" & [Segment2Desc] & \".\" & [Segment3Desc] A"
    "S ProductTypeView, qryProductCategorySegment2.tblProductCategorySegment1.Active,"
    " qryProductCategorySegment2.tblProductCategorySegment2.Active\015\012FROM qryPro"
    "ductCategorySegment2 INNER JOIN tblProductCategorySegment3 ON (qryProductCategor"
    "ySegment2.Segment2ID = tblProductCategorySegment3.Segment2ID) AND (qryProductCat"
    "egorySegment2.Segment1ID = tblProductCategorySegment3.Segment1ID)\015\012WHERE ("
    "((qryProductCategorySegment2.tblProductCategorySegment1.Active)=-1) AND ((qryPro"
    "ductCategorySegment2.tblProductCategorySegment2.Active)=-1) AND ((tblProductCate"
    "gorySegment3.Active)=-1))\015\012ORDER BY tblProductCategorySegment3.Segment1ID,"
    " tblProductCategorySegment3.Segment2ID, tblProductCategorySegment3.Segment3ID;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbInteger "RowHeight" ="240"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xaac754695b3afe4ba56fea8369e0a61b
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
            0x7ae258b4b07650498093439a642dda11
        End
    End
    Begin
        dbText "Name" ="ProductTypeView"
        dbInteger "ColumnWidth" ="4095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1965b8c16ab81d4680d6a840adaae123
        End
    End
    Begin
        dbText "Name" ="tblProductCategorySegment3.Segment3Desc"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryProductCategorySegment2.Segment1Desc"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryProductCategorySegment2.Segment2Desc"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
