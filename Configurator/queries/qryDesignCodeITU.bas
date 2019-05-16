Operation =1
Option =0
Having ="(((tbl_Fibers.ITU) Is Not Null) AND ((tbl_Fibers.Active)=True))"
Begin InputTables
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Expression ="tbl_Fibers.ITU"
    Expression ="tbl_Fibers.Active"
End
Begin Groups
    Expression ="tbl_Fibers.ITU"
    GroupLevel =0
    Expression ="tbl_Fibers.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7c4f51b5fd1bce479a8cf648ee2770f9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =242
    Top =185
    Right =1898
    Bottom =566
    Left =-1
    Top =-1
    Right =1624
    Bottom =150
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =327
        Bottom =169
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
