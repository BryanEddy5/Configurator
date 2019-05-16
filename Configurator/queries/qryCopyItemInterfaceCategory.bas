Operation =3
Name ="tblItemCategoryInterface"
Option =0
Where ="(((tblItemInterfaceDataTypes.AppliesTo)=\"Category\") And ((tblItemInterfaceData"
    "Types.DesignType)=10) And ((tblItemInterfaceDataTypes.ItemType)=Forms!frmMainMen"
    "u!NewItemType))"
Begin InputTables
    Name ="tblItemInterfaceDataTypes"
    Name ="qryItemInterfaceCategoryExisting"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="ItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="tblItemInterfaceDataTypes.DataTypeOrg"
    Name ="CategorySetName"
    Expression ="tblItemInterfaceDataTypes.DataTypeName"
    Alias ="CategoryNameToLoad"
    Name ="CategoryName"
    Expression ="IIf([CategoryName] Is Not Null,[CategoryName],[DataTypeValue])"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
    Name ="AllowOverride"
    Expression ="tblItemInterfaceDataTypes.AllowOverride"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
End
Begin Joins
    LeftTable ="tblItemInterfaceDataTypes"
    RightTable ="qryItemInterfaceCategoryExisting"
    Expression ="tblItemInterfaceDataTypes.DataTypeOrg = qryItemInterfaceCategoryExisting.Organiz"
        "ationCode"
    Flag =2
    LeftTable ="tblItemInterfaceDataTypes"
    RightTable ="qryItemInterfaceCategoryExisting"
    Expression ="tblItemInterfaceDataTypes.DataTypeName = qryItemInterfaceCategoryExisting.Catego"
        "rySetName"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc878cc9487412f45b5ad0bae71bb015d
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf0c1015982820d4aaeca849a8fd6897d
        End
    End
    Begin
        dbText "Name" ="CategoryNameToLoad"
        dbInteger "ColumnWidth" ="3315"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xce12ba97a9801a45b0c45861a5caedc6
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x958ddd294d13f14eb7ae40444a7f357d
        End
    End
End
Begin
    State =0
    Left =147
    Top =127
    Right =1418
    Bottom =614
    Left =-1
    Top =-1
    Right =1239
    Bottom =222
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =182
        Top =23
        Right =450
        Bottom =246
        Top =0
        Name ="tblItemInterfaceDataTypes"
        Name =""
    End
    Begin
        Left =680
        Top =38
        Right =879
        Bottom =156
        Top =0
        Name ="qryItemInterfaceCategoryExisting"
        Name =""
    End
End
