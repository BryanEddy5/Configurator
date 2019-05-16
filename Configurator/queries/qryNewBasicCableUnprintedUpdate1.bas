Operation =4
Option =0
Begin InputTables
    Name ="tblNewBasicCable"
End
Begin OutputColumns
    Name ="tblNewBasicCable.[New Oracle Part Number]"
    Expression ="Left([New Oracle Part Number],7) & \"-00\""
    Name ="tblNewBasicCable.[Item No]"
    Expression ="Left([Item No],8) & \"U\" & Mid([Item No],10,20)"
    Name ="tblNewBasicCable.Customer"
    Expression ="\"UNPRINTED\""
    Name ="tblNewBasicCable.[Customer Part Num]"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Reel No]"
    Expression ="0"
    Name ="tblNewBasicCable.[Print Item No]"
    Expression ="0"
    Name ="tblNewBasicCable.[Print Type (base)]"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Line 1]"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Line 2]"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Line 3]"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Line 4]"
    Expression ="\"\""
    Name ="tblNewBasicCable.NewPrintLine4"
    Expression ="\"\""
    Name ="tblNewBasicCable.PrintSpacing"
    Expression ="\"\""
    Name ="tblNewBasicCable.PrintNotes"
    Expression ="\"\""
    Name ="tblNewBasicCable.[Print Height]"
    Expression ="\"\""
    Name ="tblNewBasicCable.CustomerRev"
    Expression ="Null"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x830e235e062fba498ae87b207ed39cf7
End
Begin
    Begin
        dbText "Name" ="tblNewBasicCable.[New Oracle Part Number]"
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
    Bottom =323
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =318
        Top =6
        Right =560
        Bottom =349
        Top =0
        Name ="tblNewBasicCable"
        Name =""
    End
End
