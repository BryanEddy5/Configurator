dbMemo "SQL" ="INSERT INTO CutSheet_BaseImages ( Color, ImageID, ImageTypeID, Base )\015\012SEL"
    "ECT DISTINCT tblColors_Available.Color, CutSheet_Images.ImageID, CutSheet_Images"
    ".ImageTypeID, tblFoundImages.Base\015\012FROM (tblFoundImages INNER JOIN tblColo"
    "rs_Available ON tblFoundImages.Color = tblColors_Available.Color_ID) INNER JOIN "
    "CutSheet_Images ON (CutSheet_Images.ImageTypeID = tblFoundImages.ImageType) AND "
    "(tblFoundImages.ImageName = CutSheet_Images.ImageName) AND (tblColors_Available."
    "Color = CutSheet_Images.Color);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x46da670b29ea714caca32b35eddbb4de
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="tblColors_Available.Color"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.ImageName"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.ImageType"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.Location"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.Base"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.Color"
        dbInteger "ColumnWidth" ="3105"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CutSheet_Images.ImageID"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CutSheet_Images.ImageTypeID"
        dbLong "AggregateType" ="-1"
    End
End
