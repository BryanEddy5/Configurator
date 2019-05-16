Operation =1
Option =16
RowCount ="2,"
Having ="(((tblColors_Available.Color_ID)=Right(Left([Forms]![frmCreateItem]![ItemNo],10)"
    ",1)))"
Begin InputTables
    Name ="tblColors_Available"
End
Begin OutputColumns
    Expression ="tblColors_Available.Color"
End
Begin Groups
    Expression ="tblColors_Available.Color"
    GroupLevel =0
    Expression ="tblColors_Available.Color_ID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe8feb80a9a001f489a38f49dbb759697
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblColors_Available.Color"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =116
    Right =1605
    Bottom =941
    Left =-1
    Top =-1
    Right =1454
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =368
        Top =39
        Right =680
        Bottom =322
        Top =0
        Name ="tblColors_Available"
        Name =""
    End
End
