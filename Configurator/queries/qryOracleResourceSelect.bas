Operation =1
Option =0
Where ="(((tblOracleResourceCodes.DepartmentCode)=[Forms]![frmMainMenu]![OracleDepartmen"
    "tCode]))"
Begin InputTables
    Name ="tblOracleResourceCodes"
End
Begin OutputColumns
    Expression ="tblOracleResourceCodes.ResourceCode"
    Expression ="tblOracleResourceCodes.ResourceCodeDescription"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x89707e3c890d1a4fb025d980c3751a6c
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
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tblOracleResourceCodes"
        Name =""
    End
End
