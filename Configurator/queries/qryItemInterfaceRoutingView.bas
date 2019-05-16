Operation =1
Option =0
Where ="(((tblBOMOpSequencesInterface.AssemblyItemNumber)=[Forms]![frmItemInterface]![It"
    "emNumber]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Expression ="tblBOMOpSequencesInterface.*"
End
Begin OrderBy
    Expression ="tblBOMOpSequencesInterface.CopyCount"
    Flag =0
    Expression ="tblBOMOpSequencesInterface.OperationSeqNum"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6bb59560dc8fea49af883fc18ef45b16
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
        Right =257
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End
