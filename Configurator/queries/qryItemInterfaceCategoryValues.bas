Operation =1
Option =0
Where ="(((tblItemInterfaceCategoryValues_TEST.DesignTypeID)=10) And ((tblItemInterfaceC"
    "ategoryValues_TEST.CategorySetName)=Forms!frmItemInterface!frmItemInterfaceCateg"
    "ory.Form!CategorySetName))"
Begin InputTables
    Name ="tblItemInterfaceCategoryValues_TEST"
End
Begin OutputColumns
    Expression ="tblItemInterfaceCategoryValues_TEST.CategoryValue"
    Expression ="tblItemInterfaceCategoryValues_TEST.CategoryValueDesc"
    Expression ="tblItemInterfaceCategoryValues_TEST.DesignTypeID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3290e792393065429b9f99e2ab9a1118
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblItemInterfaceCategoryValues_TEST"
        Name =""
    End
End
