Operation =1
Option =0
Where ="(((Left([New Oracle Part #],7)) Like Forms!frmPremiseTBReplace!cboBaseSelect & \""
    "*\") And ((qryPremiseTBItems.[Fiber Oracle item])=Forms!frmPremiseTBReplace!Fibe"
    "rType))"
Begin InputTables
    Name ="qryPremiseTBItems"
    Name ="qryPremiseTBItemsCopied"
End
Begin OutputColumns
    Expression ="qryPremiseTBItemsCopied.ComponentItemNumber"
    Alias ="TBTypeCopied"
    Expression ="Mid(qryPremiseTBItemsCopied![Item No],12,1)"
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    Expression ="qryPremiseTBItems.[Jacket Material]"
    Alias ="NewComponentItem"
    Expression ="qryPremiseTBItems.[New Oracle Part #]"
    Alias ="TBTypeNew"
    Expression ="Mid(qryPremiseTBItems![Item No],12,1)"
    Alias ="NewBase"
    Expression ="Left([New Oracle Part #],7)"
    Expression ="qryPremiseTBItemsCopied.[Fiber Oracle item]"
    Expression ="qryPremiseTBItemsCopied.[Jacket Material]"
    Expression ="qryPremiseTBItems.[Item No]"
    Expression ="qryPremiseTBItemsCopied.[Item No]"
    Alias ="TB"
    Expression ="Mid(qryPremiseTBItemsCopied.[Item No],12,1)"
    Alias ="Expr1"
    Expression ="Mid(qryPremiseTBItems.[Item No],7,2)"
End
Begin Joins
    LeftTable ="qryPremiseTBItems"
    RightTable ="qryPremiseTBItemsCopied"
    Expression ="qryPremiseTBItems.[Jacket Material] = qryPremiseTBItemsCopied.[Jacket Material]"
    Flag =1
    LeftTable ="qryPremiseTBItems"
    RightTable ="qryPremiseTBItemsCopied"
    Expression ="qryPremiseTBItems.ItemNumberSuffix = qryPremiseTBItemsCopied.ItemNumberSuffix"
    Flag =1
End
Begin OrderBy
    Expression ="qryPremiseTBItems.[Item No]"
    Flag =0
    Expression ="qryPremiseTBItemsCopied.ItemNumberSuffix"
    Flag =0
    Expression ="qryPremiseTBItems.[Fiber Oracle item]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3ac31aee8f2f59418f9230166e86deb4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TBTypeCopied"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09c32665b47b89478555802a61138520
        End
    End
    Begin
        dbText "Name" ="TBTypeNew"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4cd8d0f74ec2034b9fbbc6f62d67afd1
        End
    End
    Begin
        dbText "Name" ="NewBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x35278ed4f9cc064fa6bf112f4a8dcbc0
        End
    End
    Begin
        dbText "Name" ="TB"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1ccbd91f28d3a84082e1b7c0c023fd18
        End
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewComponentItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a6c35542c4f5741ba9d4e3c6e6359d4
        End
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItems.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseTBItemsCopied.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc1281e50ab8f354492ac165aae3ac1b2
        End
    End
End
Begin
    State =0
    Left =64
    Top =83
    Right =1173
    Bottom =530
    Left =-1
    Top =-1
    Right =1077
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =729
        Top =3
        Right =978
        Bottom =184
        Top =0
        Name ="qryPremiseTBItems"
        Name =""
    End
    Begin
        Left =128
        Top =15
        Right =475
        Bottom =163
        Top =0
        Name ="qryPremiseTBItemsCopied"
        Name =""
    End
End
