﻿Operation =1
Option =0
Where ="(((tblfrmPrintType.Number)=1 Or (tblfrmPrintType.Number)=2))"
Begin InputTables
    Name ="tblfrmPrintType"
End
Begin OutputColumns
    Expression ="tblfrmPrintType.Number"
    Expression ="tblfrmPrintType.[Print Type]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2728396ea575e24bafe5b0e3566da45a
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
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =200
        Bottom =94
        Top =0
        Name ="tblfrmPrintType"
        Name =""
    End
End
