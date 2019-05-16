Operation =1
Option =0
Begin InputTables
    Name ="tblColors_Available"
End
Begin OutputColumns
    Expression ="tblColors_Available.Color"
    Expression ="tblColors_Available.Color_ID"
    Expression ="tblColors_Available.ID"
End
Begin Groups
    Expression ="tblColors_Available.Color"
    GroupLevel =0
    Expression ="tblColors_Available.Color_ID"
    GroupLevel =0
    Expression ="tblColors_Available.ID"
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
    Begin
        dbText "Name" ="tblColors_Available.Color_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblColors_Available.ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =26
    Top =200
    Right =1264
    Bottom =852
    Left =-1
    Top =-1
    Right =1206
    Bottom =287
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =306
        Top =49
        Right =507
        Bottom =237
        Top =0
        Name ="tblColors_Available"
        Name =""
    End
End
