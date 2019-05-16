Operation =1
Option =2
Where ="(((tblItemInterfaceDataTypes.DesignType)=10))"
Begin InputTables
    Name ="tblItemInterfaceDataTypes"
End
Begin OutputColumns
    Expression ="tblItemInterfaceDataTypes.ItemType"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xab9a06f2f93c7940b10043058b1ca9f2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ItemType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x88f1a59a6a043f418879ebef76067919
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tblItemInterfaceDataTypes"
        Name =""
    End
End
