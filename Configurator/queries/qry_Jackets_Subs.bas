Operation =1
Option =0
Having ="((([Basic Product Construction].Active)=True) AND (([Basic Product Construction]"
    ".[New Oracle Part #]) Not Like \"*test*\") AND (([Basic Product Construction].[I"
    "tem No]) Not Like \"*bif*\" And ([Basic Product Construction].[Item No]) Not Lik"
    "e \"*crg*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Alias ="Jacket_Material_letter"
    Expression ="Mid([Item No],2,1)"
    Alias ="Jacket_Color_letter"
    Expression ="Mid([Item No],10,1)"
    Expression ="[Basic Product Construction].Base"
    Alias ="AddOn"
    Expression ="Mid([Basic Product Construction]![Item No],14,7)"
    Alias ="JacketBase"
    Expression ="Left([Jacket Material],8)"
    Alias ="CountOfJacket Color"
    Expression ="Count([Basic Product Construction].[Jacket Color])"
End
Begin OrderBy
    Expression ="Mid([Item No],2,1)"
    Flag =0
    Expression ="Mid([Item No],10,1)"
    Flag =0
End
Begin Groups
    Expression ="[Basic Product Construction].[Jacket Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Color Chip ID]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Jacket Color]"
    GroupLevel =0
    Expression ="Mid([Item No],2,1)"
    GroupLevel =0
    Expression ="Mid([Item No],10,1)"
    GroupLevel =0
    Expression ="[Basic Product Construction].Base"
    GroupLevel =0
    Expression ="Mid([Basic Product Construction]![Item No],14,7)"
    GroupLevel =0
    Expression ="Left([Jacket Material],8)"
    GroupLevel =0
    Expression ="[Basic Product Construction].Active"
    GroupLevel =0
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Item No]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbb645f67f805ca4786ac13079c66cef9
End
dbInteger "RowHeight" ="285"
dbMemo "Filter" ="((([qry_Jackets_Subs].[Jacket Material] Like \"*130*\"))) AND ([qry_Jackets_Subs"
    "].[Jacket_Color_letter]=\"7\")"
dbMemo "OrderBy" ="[qry_Jackets_Subs].[Jacket_Color_letter], [qry_Jackets_Subs].[Jacket_Material_le"
    "tter]"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbInteger "ColumnWidth" ="4275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jacket_Material_letter"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x11dc051f84a3b64da894a1e28c50fe84
        End
    End
    Begin
        dbText "Name" ="Jacket_Color_letter"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6498340509ae954198612ff8eaf29367
        End
    End
    Begin
        dbText "Name" ="AddOn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb0cc8fd425dc2b4e9c222de6965ace5c
        End
    End
    Begin
        dbText "Name" ="JacketBase"
        dbInteger "ColumnWidth" ="4275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5d9518041d2e244692960f47a9a9aaec
        End
    End
    Begin
        dbText "Name" ="CountOfJacket Color"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf3b34275b214024cb918483b800ef482
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jacket Material"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc915cf4a8de2ab4db909d452e5d3ef32
        End
    End
    Begin
        dbText "Name" ="Color Chip ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2f4bcf88d0468b4686cde1e3dcf6b745
        End
    End
    Begin
        dbText "Name" ="Jacket Color"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc9ca8e9c4afcd409e4fda393392c364
        End
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x696d81c4e3c1104d9a0ee28e95b0e590
        End
    End
End
Begin
    State =0
    Left =-786
    Top =622
    Right =941
    Bottom =1353
    Left =-1
    Top =-1
    Right =1695
    Bottom =342
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =401
        Bottom =319
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End
