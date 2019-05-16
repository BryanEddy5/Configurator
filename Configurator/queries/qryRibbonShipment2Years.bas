Operation =1
Option =0
Where ="(((tblShipmentsCable.item_number) Like \"pw*\" Or (tblShipmentsCable.item_number"
    ") Like \"rbn*\") And ((tblShipmentsCable.transaction_date)>#9/28/2013#))"
Begin InputTables
    Name ="tblShipmentsCable"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblShipmentsCable.item_number"
    Alias ="Expr2"
    Expression ="tblShipmentsCable.bill_to_name"
    Alias ="Expr3"
    Expression ="tblShipmentsCable.transaction_date"
    Alias ="Qty"
    Expression ="Int([quantity])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x92650a28340f9b4788e7539915929864
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="4770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c275f9d57c18b4e9170d134aacd22d7
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe5cc6b136e87fa49ada1a184bef2f6c0
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c6c9a710cff314b8ecf85758655a812
        End
    End
    Begin
        dbText "Name" ="Qty"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdc2219d42e49c441b05f51b68b189910
        End
    End
End
Begin
    State =0
    Left =101
    Top =112
    Right =1210
    Bottom =715
    Left =-1
    Top =-1
    Right =1077
    Bottom =227
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =145
        Top =1
        Right =343
        Bottom =170
        Top =0
        Name ="tblShipmentsCable"
        Name =""
    End
End
