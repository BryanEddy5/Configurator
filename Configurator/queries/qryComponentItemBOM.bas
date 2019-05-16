Operation =1
Option =0
Where ="((([tblAFLPRD_BOMInvComp_CAB].[AssemblyItemNumber])=[Forms]![frmMainMenu]![Compo"
    "nentItemNumber]))"
Begin InputTables
    Name ="tblAFLPRD_BOMInvComp_CAB"
End
Begin OutputColumns
    Expression ="tblAFLPRD_BOMInvComp_CAB.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfe1e0d1b3ffd264e851479c17f1ae974
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
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =0
        Name ="tblAFLPRD_BOMInvComp_CAB"
        Name =""
    End
End
