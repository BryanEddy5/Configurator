Operation =1
Option =0
Where ="(((tblQAItemSpecElements.SpecName)=[Forms]![frmItemInterface]![ItemNumber]))"
Begin InputTables
    Name ="tblQAItemSpecElements"
End
Begin OutputColumns
    Expression ="tblQAItemSpecElements.*"
End
Begin Parameters
    Name ="[Forms]![frmItemInterface]![ItemNumber]"
    Flag =10
End
Begin OrderBy
    Expression ="tblQAItemSpecElements.ElementName"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x210443fdeea53041b51c5fd6559809f1
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
    Right =1418
    Bottom =518
    Left =-1
    Top =-1
    Right =1368
    Bottom =264
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =111
        Top =20
        Right =383
        Bottom =228
        Top =0
        Name ="tblQAItemSpecElements"
        Name =""
    End
End
