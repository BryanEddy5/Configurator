Operation =3
Name ="tblNewPrductConstruction"
Option =0
Begin InputTables
    Name ="qryRDBasicConstruction"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].*"
End
Begin Joins
    LeftTable ="qryRDBasicConstruction"
    RightTable ="Basic Product Construction"
    Expression ="qryRDBasicConstruction.[New Oracle Part #] = [Basic Product Construction].[New O"
        "racle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6f236729113d764d9389784b41ff7775
End
dbBoolean "UseTransaction" ="-1"
Begin
End
Begin
    State =0
    Left =147
    Top =173
    Right =1439
    Bottom =856
    Left =-1
    Top =-1
    Right =1260
    Bottom =452
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =470
        Top =6
        Right =566
        Bottom =124
        Top =0
        Name ="qryRDBasicConstruction"
        Name =""
    End
    Begin
        Left =57
        Top =13
        Right =432
        Bottom =386
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
