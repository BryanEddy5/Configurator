Operation =4
Option =0
Where ="(((Left([ComponentItemNumber],7))=Left([Forms]![frmPremiseTBReplace]![ComponentI"
    "temNumber],7)))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Left(Forms!frmPremiseTBReplace!NewComponentItem,7) & Right(Forms!frmPremiseTBRep"
        "lace!NewComponentItem,2)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0xe4b194497b5f4d41bc2fea88c23d2b0d
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
        Right =251
        Bottom =124
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End
