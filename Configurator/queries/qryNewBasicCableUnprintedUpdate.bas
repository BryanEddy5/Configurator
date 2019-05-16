Operation =4
Option =0
Begin InputTables
    Name ="tblNewBasicCable_New"
End
Begin OutputColumns
    Name ="tblNewBasicCable_New.[New Oracle Part #]"
    Expression ="Left([New Oracle Part #],7) & \"-00\""
    Name ="tblNewBasicCable_New.[Item No]"
    Expression ="Left([Item No],8) & \"U\" & Mid([Item No],10,20)"
    Name ="tblNewBasicCable_New.Customer"
    Expression ="\"UNPRINTED\""
    Name ="tblNewBasicCable_New.[Print Item No]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Reel No]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Type (base)]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Height]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Line 1]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Line 2]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Line 3]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Line 4]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.NewPrintLine4"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Print Spacing]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.PrintNotes"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Customer Part#]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.Active"
    Expression ="-1"
    Name ="tblNewBasicCable_New.CustomerRev"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Special Instr Product1]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.[zzSpecial Instr Product4]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.ENumber"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Listing Company]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[UL-ETL-Listing]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[UL-ETL-Const]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[UL-ETL-Section]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.[Label Type]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.VendorSpecificFiber"
    Expression ="GetisVendorSpecificPrint()"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xea6f4b4598c01a4fa8fe0616740e7aaa
End
Begin
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Reel No]"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Type (base)]"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Height]"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 1]"
        dbInteger "ColumnWidth" ="4050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 2]"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 3]"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 4]"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.NewPrintLine4"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Item No]"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Customer"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Customer Part Num]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Item No]"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Type (base)]"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 3]"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Customer"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Height]"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 4]"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Customer Part Num]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 1]"
        dbInteger "ColumnWidth" ="4050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.NewPrintLine4"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Reel No]"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 2]"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Customer Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[zzSpecial Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.VendorSpecificFiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =329
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =468
        Top =12
        Right =826
        Bottom =244
        Top =0
        Name ="tblNewBasicCable_New"
        Name =""
    End
End
