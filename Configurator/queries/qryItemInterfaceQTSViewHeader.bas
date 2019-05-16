Operation =1
Option =0
Where ="(((tblQAItemSpecs.SpecName)=[Forms]![frmItemInterface]![ItemNumber]))"
Begin InputTables
    Name ="tblQAItemSpecs"
End
Begin OutputColumns
    Expression ="tblQAItemSpecs.*"
End
Begin Parameters
    Name ="[Forms]![frmItemInterface]![ItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb6a592473e036f438ba82cfaf124dcb7
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
        Left =421
        Top =6
        Right =517
        Bottom =124
        Top =0
        Name ="tblQAItemSpecs"
        Name =""
    End
End
