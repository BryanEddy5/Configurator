Operation =4
Option =0
Begin InputTables
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[New Oracle Part #]"
    Expression ="IIf([New Oracle Part #] Like \"rbn*\",[New Oracle Part #],Left([New Oracle Part "
        "#],7) & \"-00\")"
    Name ="tblNewPrductConstruction.[Item No]"
    Expression ="UpdateToUnprintedCatalog([Item No])"
    Name ="tblNewPrductConstruction.Customer"
    Expression ="\"UNPRINTED\""
    Name ="tblNewPrductConstruction.[Customer Part#]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Reel No]"
    Expression ="0"
    Name ="tblNewPrductConstruction.[Print Item No]"
    Expression ="0"
    Name ="tblNewPrductConstruction.[Print Type (base)]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Line 1]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Line 2]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Line 3]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Line 4]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.NewPrintLine4"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Print Spacing]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.PrintNotes"
    Expression ="Null"
    Name ="tblNewPrductConstruction.Active"
    Expression ="-1"
    Name ="tblNewPrductConstruction.CustomerRev"
    Expression ="Null"
    Name ="tblNewPrductConstruction.ENumber"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Label Type]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[UL-ETL-Listing]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[UL-ETL-Const]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[UL-ETL-Section]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Listing Company]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.VendorSpecificFiber"
    Expression ="GetisVendorSpecificPrint()"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x49ae324642d87a43845041ee8a6193d3
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Type (base)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Item No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3135"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Customer Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.VendorSpecificFiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1197
    Bottom =775
    Left =-1
    Top =-1
    Right =1077
    Bottom =265
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =290
        Bottom =292
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End
