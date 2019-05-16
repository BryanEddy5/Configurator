Operation =1
Option =0
Where ="(((tblQASpecsFiberDefaults.FiberType)=[Forms]![frmItemInterfaceQTSFiberSelect]!["
    "cboFiberSelect]))"
Begin InputTables
    Name ="tblQASpecsFiberDefaults"
End
Begin OutputColumns
    Expression ="tblQASpecsFiberDefaults.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7f40508f791759449b1329c277683cb6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =129
    Top =127
    Right =1238
    Bottom =589
    Left =-1
    Top =-1
    Right =1077
    Bottom =248
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =249
        Bottom =169
        Top =0
        Name ="tblQASpecsFiberDefaults"
        Name =""
    End
End
