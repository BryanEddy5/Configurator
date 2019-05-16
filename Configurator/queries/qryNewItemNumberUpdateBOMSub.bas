Operation =4
Option =0
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber)=[Forms]![frmItemInterface]"
    "![frmPremiseItemDetailsCopiedSubunit].[Form]![New Oracle Part #]))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetailsSubunit.Form![New Oracle Part #]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc871722c5c8cb34ebc3b14223d03f84c
End
Begin
End
Begin
    State =0
    Left =3
    Top =0
    Right =1112
    Bottom =324
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =308
        Top =6
        Right =526
        Bottom =124
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End
