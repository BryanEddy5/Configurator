dbMemo "SQL" ="INSERT INTO CutSheet_Images ( ImageName, ImageTypeID, Color )\015\012SELECT DIST"
    "INCT tblFoundImages.ImageName, tblFoundImages.ImageType, tblColors_Available.Col"
    "or\015\012FROM tblFoundImages INNER JOIN tblColors_Available ON tblFoundImages.C"
    "olor = tblColors_Available.Color_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x3eff9a0ded4e3748a563c2e4c8fffb47
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblFoundImages.ImageName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.ImageType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFoundImages.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblColors_Available.Color"
        dbLong "AggregateType" ="-1"
    End
End
