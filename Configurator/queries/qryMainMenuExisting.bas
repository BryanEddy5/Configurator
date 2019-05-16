Operation =1
Option =2
Begin InputTables
    Name ="qryItemNoSearch"
    Name ="qryMainMenuMatch"
End
Begin OutputColumns
    Expression ="qryMainMenuMatch.Oracle"
    Alias ="Exists"
    Expression ="IIf(qryMainMenuMatch![Item No] Is Not Null,-1,0)"
    Expression ="qryMainMenuMatch.[Item No]"
    Expression ="qryItemNoSearch.Active"
    Expression ="qryItemNoSearch.Reason"
End
Begin Joins
    LeftTable ="qryItemNoSearch"
    RightTable ="qryMainMenuMatch"
    Expression ="qryItemNoSearch.[Item No] = qryMainMenuMatch.[Item No]"
    Flag =2
End
Begin OrderBy
    Expression ="qryMainMenuMatch.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0a51ebbace02774ab75299fdbca88111
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Exists"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1162e1015760424f99c25a0f1273f1af
        End
    End
    Begin
        dbText "Name" ="qryMainMenuMatch.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMainMenuMatch.[Item No]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =40
    Top =108
    Right =1332
    Bottom =545
    Left =-1
    Top =-1
    Right =1260
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =396
        Bottom =212
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
    Begin
        Left =519
        Top =27
        Right =811
        Bottom =185
        Top =0
        Name ="qryMainMenuMatch"
        Name =""
    End
End
