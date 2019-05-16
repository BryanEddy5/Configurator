Operation =4
Option =0
Where ="(((tblBOMOpResourcesInterface.ResourceCode)<>\"Setup\") And ((tblBOMOpResourcesI"
    "nterface.OperationSeqNumber)=Forms!frmMainMenu!OperationSequence))"
Begin InputTables
    Name ="tblBOMOpResourcesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpResourcesInterface.ResourceCode"
    Expression ="\"\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd289c2647ba7534e84bc0efdb7e4c5a6
End
Begin
End
Begin
    State =0
    Left =62
    Top =45
    Right =1375
    Bottom =502
    Left =-1
    Top =-1
    Right =1281
    Bottom =247
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =332
        Bottom =255
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
End
