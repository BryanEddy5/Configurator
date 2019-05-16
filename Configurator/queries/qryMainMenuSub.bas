Operation =1
Option =0
Where ="(((qryMainMenuExisting.Active)<>0))"
Begin InputTables
    Name ="qryMainMenuExisting"
End
Begin OutputColumns
    Expression ="qryMainMenuExisting.Oracle"
    Expression ="qryMainMenuExisting.[Exists]"
    Expression ="qryMainMenuExisting.[Item No]"
    Expression ="qryMainMenuExisting.Active"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x75a1fa5655712548b7c6a21a3a88c560
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryMainMenuExisting.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMainMenuExisting.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMainMenuExisting.[Exists]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMainMenuExisting.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exists"
        dbBinary "GUID" = Begin
            0xb9229f0e504bd441bfe9d9152c5679a3
        End
    End
End
Begin
    State =0
    Left =-188
    Top =343
    Right =1104
    Bottom =765
    Left =-1
    Top =-1
    Right =1260
    Bottom =174
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =236
        Bottom =154
        Top =0
        Name ="qryMainMenuExisting"
        Name =""
    End
End
