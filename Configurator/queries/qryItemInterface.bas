Operation =1
Option =0
Begin InputTables
    Name ="tblItemInterface"
End
Begin OutputColumns
    Expression ="tblItemInterface.*"
End
Begin OrderBy
    Expression ="tblItemInterface.ItemNumber"
    Flag =0
    Expression ="tblItemInterface.OrganizationCode"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb99a38401c6e6b4eb348033aa1a83602
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =40
    Right =1652
    Bottom =364
    Left =-1
    Top =-1
    Right =1602
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =271
        Bottom =109
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
End
