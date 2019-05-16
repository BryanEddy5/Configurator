Operation =3
Name ="tblQAItemSpecElements"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="Listing"
    Name ="ElementName"
    Expression ="\"CABLE LISTING\""
    Alias ="SpecName"
    Name ="SpecName"
    Expression ="[Forms]![frmItemInterface]![ItemNumber]"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x6075d87fedb0c044acbf2b900e28de7f
End
Begin
    Begin
        dbText "Name" ="Listing"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5acbdbe775769e4e9a6de00a4040b6e8
        End
    End
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x737660c66e274c47a5226f8018fcd2e0
        End
    End
End
Begin
    State =0
    Left =321
    Top =179
    Right =1613
    Bottom =641
    Left =-1
    Top =-1
    Right =1260
    Bottom =248
    Left =0
    Top =0
    ColumnsShown =651
End
