Operation =1
Option =0
Where ="((([tblAFLPRD_BOMOpSeq_CAB].[AssemblyItemNumber])=[Forms]![frmMainMenu]![Compone"
    "ntItemNumber]))"
Begin InputTables
    Name ="tblAFLPRD_BOMOpSeq_CAB"
End
Begin OutputColumns
    Expression ="tblAFLPRD_BOMOpSeq_CAB.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3185a6ff06df8d4ab9a3cb9f2d5c3c2b
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
        Left =38
        Top =6
        Right =227
        Bottom =124
        Top =0
        Name ="tblAFLPRD_BOMOpSeq_CAB"
        Name =""
    End
End
