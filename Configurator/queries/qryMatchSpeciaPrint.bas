Operation =1
Option =16
RowCount ="1"
Having ="(((qrySpecialInstructions.Base)=[Forms]![frmCreateItem]![Base]))"
Begin InputTables
    Name ="qrySpecialInstructions"
End
Begin OutputColumns
    Expression ="qrySpecialInstructions.Base"
    Expression ="qrySpecialInstructions.Special_1"
    Expression ="qrySpecialInstructions.Special_2"
    Expression ="qrySpecialInstructions.Special_3"
    Expression ="qrySpecialInstructions.Special_4"
    Alias ="CountOfSpecial_1"
    Expression ="Count(qrySpecialInstructions.Special_1)"
End
Begin OrderBy
    Expression ="Count(qrySpecialInstructions.Special_1)"
    Flag =1
End
Begin Groups
    Expression ="qrySpecialInstructions.Base"
    GroupLevel =0
    Expression ="qrySpecialInstructions.Special_1"
    GroupLevel =0
    Expression ="qrySpecialInstructions.Special_2"
    GroupLevel =0
    Expression ="qrySpecialInstructions.Special_3"
    GroupLevel =0
    Expression ="qrySpecialInstructions.Special_4"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xefecf477afa3344e931909137cc1493c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfSpecial_1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xad95be945007b448972d63d1a1721296
        End
    End
    Begin
        dbText "Name" ="Special_2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x233dec7d89b90a4f89d2221e6c0570ff
        End
    End
End
Begin
    State =0
    Left =91
    Top =199
    Right =1383
    Bottom =611
    Left =-1
    Top =-1
    Right =1260
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =236
        Bottom =169
        Top =0
        Name ="qrySpecialInstructions"
        Name =""
    End
End
