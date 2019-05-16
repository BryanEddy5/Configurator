Operation =1
Option =2
Begin InputTables
    Name ="tblOracleItems"
End
Begin OutputColumns
    Expression ="tblOracleItems.ItemNumber"
    Alias ="ViewOrder"
    Expression ="IIf(Left([ItemNumber],2)=Left(Forms!frmMainMenu!NewItemNumber,2),-1,0)"
    Expression ="tblOracleItems.Description"
End
Begin OrderBy
    Expression ="IIf(Left([ItemNumber],2)=Left(Forms!frmMainMenu!NewItemNumber,2),-1,0)"
    Flag =0
    Expression ="tblOracleItems.ItemNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf10de4fa675c7e4fb4c6de19a6cec711
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbMemo "Caption" ="Oracle Item Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ViewOrder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x394ed7a4f5529c45a5a9ab2867310735
        End
    End
    Begin
        dbText "Name" ="tblOracleItems.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =122
    Top =315
    Right =1231
    Bottom =837
    Left =-1
    Top =-1
    Right =1077
    Bottom =274
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =68
        Top =26
        Right =303
        Bottom =172
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End
