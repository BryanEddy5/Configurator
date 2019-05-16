Operation =4
Option =0
Begin InputTables
    Name ="tblBOMBillOfMtlsInterface"
End
Begin OutputColumns
    Name ="tblBOMBillOfMtlsInterface.ItemNumber"
    Expression ="Forms!frmMainMenu!NewPrintExtension"
    Name ="tblBOMBillOfMtlsInterface.CommonItemNumber"
    Expression ="Forms!frmMainMenu!NewBaseItem"
    Name ="tblBOMBillOfMtlsInterface.CommonOrgCode"
    Expression ="tblBOMBillOfMtlsInterface!OrganizationCode"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb0830efe8d7c0546aeb28e7a65e29373
End
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =285
        Bottom =124
        Top =0
        Name ="tblBOMBillOfMtlsInterface"
        Name =""
    End
End
