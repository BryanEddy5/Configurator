Operation =4
Option =0
Where ="(((tblItemCategoryInterface.CategorySetName)=\"Fiber Family 1\") And ((tbl_Fiber"
    "s.Fiber)=Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType1))"
Begin InputTables
    Name ="tblItemCategoryInterface"
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="[OracleFiberFamily]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xbd4b4764f5071647bdc6b3e8121f970f
End
Begin
End
Begin
    State =0
    Left =94
    Top =136
    Right =1203
    Bottom =620
    Left =-1
    Top =-1
    Right =1077
    Bottom =270
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =339
        Top =7
        Right =562
        Bottom =170
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =244
        Bottom =244
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End
