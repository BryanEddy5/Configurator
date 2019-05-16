Operation =3
Name ="tblBOMOperationSetupParameters"
Option =0
Where ="(((tblNetworkBOMOperationSetupParameters.ItemNumber)=[Forms]![frmMainMenu]![Item"
    "NumberSelect]))"
Begin InputTables
    Name ="tblNetworkBOMOperationSetupParameters"
End
Begin OutputColumns
    Alias ="NewItemNumber"
    Name ="ItemNumber"
    Expression ="Forms!frmMainMenu!NewItemNumber"
    Name ="OrganizationCode"
    Expression ="tblNetworkBOMOperationSetupParameters.OrganizationCode"
    Name ="OperationSeqNum"
    Expression ="tblNetworkBOMOperationSetupParameters.OperationSeqNum"
    Name ="OperationCode"
    Expression ="tblNetworkBOMOperationSetupParameters.OperationCode"
    Name ="ElementName"
    Expression ="tblNetworkBOMOperationSetupParameters.ElementName"
    Name ="ElementValue"
    Expression ="tblNetworkBOMOperationSetupParameters.ElementValue"
    Alias ="CopyCount"
    Name ="CopyCount"
    Expression ="Forms!frmMainMenu!CopyCount"
    Alias ="ItemCopied"
    Name ="ItemCopied"
    Expression ="Forms!frmMainMenu!ItemNumberSelect"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![ItemNumberSelect]"
    Flag =10
    Name ="[Forms]![frmMainMenu]![CopyCount]"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd1efc0ca070f464a934b55dd81d47b0b
End
Begin
    Begin
        dbText "Name" ="NewItemNumber"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c7879ef7d4cee459c5dd4dd3b1d3c53
        End
    End
    Begin
        dbText "Name" ="CopyCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed6060e3a33e1e4298f3ae007b9ee6dc
        End
    End
    Begin
        dbText "Name" ="ItemCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1cee9aacc4f4f947ae7a213aeb748831
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1418
    Bottom =497
    Left =-1
    Top =-1
    Right =1368
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =515
        Top =6
        Right =831
        Bottom =120
        Top =0
        Name ="tblNetworkBOMOperationSetupParameters"
        Name =""
    End
End
