Operation =1
Option =2
Begin InputTables
    Name ="qryPremiseTBReplace"
End
Begin OutputColumns
    Alias ="NewBaseComponent"
    Expression ="IIf(Left([NewComponentItem],3)=\"rbn\",[NewComponentItem],Left([NewComponentItem"
        "],7))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe6240b13f40ee44b938574f63f6630cf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewBaseComponent"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x849af1b22e92044a98993c41ed90475e
        End
    End
End
Begin
    State =0
    Left =-503
    Top =275
    Right =606
    Bottom =801
    Left =-1
    Top =-1
    Right =1077
    Bottom =215
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =348
        Bottom =250
        Top =0
        Name ="qryPremiseTBReplace"
        Name =""
    End
End
