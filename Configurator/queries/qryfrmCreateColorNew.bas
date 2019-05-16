Operation =1
Option =0
Having ="((([Basic Product Construction].Active)=True) AND (([Basic Product Construction]"
    ".[New Oracle Part #]) Not Like \"*test*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Alias ="Jacket_Material_letter"
    Expression ="Right(Left([Item No],2),1)"
    Alias ="Jacket_Color_letter"
    Expression ="Right(Left([Item No],10),1)"
    Expression ="[Basic Product Construction].Active"
    Alias ="CountOfJacket Color"
    Expression ="Count([Basic Product Construction].[Jacket Color])"
End
Begin OrderBy
    Expression ="Right(Left([Item No],2),1)"
    Flag =0
    Expression ="Right(Left([Item No],10),1)"
    Flag =0
End
Begin Groups
    Expression ="[Basic Product Construction].[Jacket Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Color Chip ID]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Jacket Color]"
    GroupLevel =0
    Expression ="Right(Left([Item No],2),1)"
    GroupLevel =0
    Expression ="Right(Left([Item No],10),1)"
    GroupLevel =0
    Expression ="[Basic Product Construction].Active"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9e793a0f738f854193364d5e1ccf4516
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Jacket_Material_letter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3cfded618befa1418f270f4562ea7a42
        End
    End
    Begin
        dbText "Name" ="Jacket_Color_letter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb56a9c67e639bd4fb8fdadee7d5bc2e8
        End
    End
    Begin
        dbText "Name" ="CountOfJacket Color"
        dbInteger "ColumnWidth" ="3930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0dff16a18f217b48923b85f8d630da91
        End
    End
End
Begin
    State =0
    Left =104
    Top =241
    Right =1898
    Bottom =771
    Left =-1
    Top =-1
    Right =1762
    Bottom =299
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =470
        Bottom =319
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
