Operation =1
Option =16
RowCount ="1"
Where ="(((tblArmoredSetups.MaxCoreOD)>=CDbl([BaseOD])) And ((tblArmoredSetups.JacketTyp"
    "e)=Forms!frmCreateItem!ArmorType) And ((tblArmoredSetups.JacketMaterial)=Forms!f"
    "rmCreateItem!Jacket))"
Begin InputTables
    Name ="qryArmorODMatch"
    Name ="tblArmoredSetups"
End
Begin OutputColumns
    Expression ="tblArmoredSetups.JacketType"
    Expression ="tblArmoredSetups.MaxCoreOD"
    Expression ="tblArmoredSetups.OracleBase"
    Expression ="qryArmorODMatch.BaseOD"
    Expression ="tblArmoredSetups.MaxJacketOD"
    Expression ="tblArmoredSetups.JacketType"
    Expression ="tblArmoredSetups.ArmorNo"
    Expression ="tblArmoredSetups.JacketMaterial"
    Expression ="tblArmoredSetups.NomJacketOD"
End
Begin OrderBy
    Expression ="tblArmoredSetups.MaxJacketOD"
    Flag =0
    Expression ="tblArmoredSetups.JacketType"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x02132e7858456f40b71bb472477c5839
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =140
    Top =354
    Right =1805
    Bottom =815
    Left =-1
    Top =-1
    Right =1633
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =206
        Top =18
        Right =478
        Bottom =121
        Top =0
        Name ="qryArmorODMatch"
        Name =""
    End
    Begin
        Left =601
        Top =6
        Right =798
        Bottom =214
        Top =0
        Name ="tblArmoredSetups"
        Name =""
    End
End
