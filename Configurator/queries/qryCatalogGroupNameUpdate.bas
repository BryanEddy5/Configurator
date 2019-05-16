Operation =4
Option =0
Where ="(((tblItemInterfaceDataTypes.DesignType)=10) And ((tblItemInterfaceDataTypes.Dat"
    "aTypeName)=\"CatalogGroup\") And ((tblItemInterfaceDataTypes.ItemType)=Forms!frm"
    "MainMenu!NewItemType))"
Begin InputTables
    Name ="tblItemInterface"
    Name ="tblItemInterfaceDataTypes"
End
Begin OutputColumns
    Name ="tblItemInterface.CatalogGroupName"
    Expression ="tblItemInterfaceDataTypes!DataTypeValue"
End
Begin Joins
    LeftTable ="tblItemInterface"
    RightTable ="tblItemInterfaceDataTypes"
    Expression ="tblItemInterface.OrganizationCode = tblItemInterfaceDataTypes.DataTypeOrg"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x9b807bafcddd064c86fbe390cae48624
End
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
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =370
        Bottom =225
        Top =0
        Name ="tblItemInterfaceDataTypes"
        Name =""
    End
End
