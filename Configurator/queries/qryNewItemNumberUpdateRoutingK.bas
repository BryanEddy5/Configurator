Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterface.OperationCode) Like \"K*\"))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpSequencesInterface.OperationCode"
    Expression ="\"K\" & Format(Val(Mid(Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignC"
        "ode,3,3)),\"000\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc4524ef60e16934e9a495f3002fd41f0
End
Begin
End
Begin
    State =0
    Left =119
    Top =31
    Right =1228
    Bottom =411
    Left =-1
    Top =-1
    Right =1077
    Bottom =166
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End
