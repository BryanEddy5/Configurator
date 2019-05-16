Operation =1
Option =2
Begin InputTables
    Name ="tblCableFamily"
End
Begin OutputColumns
    Expression ="tblCableFamily.FamilyID"
    Expression ="tblCableFamily.FamilyDescription"
    Expression ="tblCableFamily.NewProductCode"
    Expression ="tblCableFamily.NewProductRelease"
    Expression ="tblCableFamily.EngineeringChange"
End
Begin OrderBy
    Expression ="tblCableFamily.FamilyID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3864924300a93c41abc45ea5e73be9e6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableFamily.NewProductRelease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableFamily.EngineeringChange"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableFamily.FamilyDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableFamily.FamilyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableFamily.NewProductCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =299
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =263
        Bottom =252
        Top =0
        Name ="tblCableFamily"
        Name =""
    End
End
