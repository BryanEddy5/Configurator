Operation =4
Option =0
Where ="(((tblItemLoadTempBOM.OperationSeqNumber)=[Forms]![frmMainMenu]![CurrentOpSeq]))"
Begin InputTables
    Name ="tblItemLoadTempBOM"
End
Begin OutputColumns
    Name ="tblItemLoadTempBOM.OperationSeqNumber"
    Expression ="Forms!frmItemInterface!frmItemInterfaceRouting.Form!OperationSeqNumber"
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
    0xd3a64124fb50414380ac164241320ec6
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
        Right =255
        Bottom =120
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
End
