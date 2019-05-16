Operation =1
Option =0
Where ="((([qryFibers].[Fiber])=[Forms]![frmCreateItem]![Fiber2]))"
Begin InputTables
    Name ="qryFibers"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryFibers.Fiber"
    Alias ="Expr2"
    Expression ="qryFibers.[OM#]"
    Alias ="Expr3"
    Expression ="qryFibers.[SM/MM]"
    Alias ="Expr4"
    Expression ="qryFibers.BIF"
    Alias ="Expr5"
    Expression ="qryFibers.ITU"
    Alias ="Expr6"
    Expression ="qryFibers.Vendor"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x19982de480dd2742b0284599d5fba044
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbBinary "GUID" = Begin
            0x5bb85c973b87bf47b4f235f01d3e1435
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbBinary "GUID" = Begin
            0x5d610ea1af8161419ee0163c109d3a95
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbBinary "GUID" = Begin
            0x79430a83654133429b2616bfd66eac9a
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbBinary "GUID" = Begin
            0x462d6be1a648f947a4cd7702e4a7f50a
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbBinary "GUID" = Begin
            0x8db50c64eb71a847b49ba825822b2b94
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbBinary "GUID" = Begin
            0xd74401990db6a2439c7bc5099e48ef1e
        End
    End
End
Begin
    State =0
    Left =361
    Top =275
    Right =1792
    Bottom =684
    Left =-1
    Top =-1
    Right =1399
    Bottom =178
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryFibers"
        Name =""
    End
End
