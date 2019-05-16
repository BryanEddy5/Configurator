Operation =1
Option =2
Where ="(((Mid([ItemNumber],8,5)) Like \"*bx*\"))"
Begin InputTables
    Name ="tblOracleItems"
    Name ="qryPremiseBase"
End
Begin OutputColumns
    Expression ="tblOracleItems.ItemNumber"
    Expression ="tblOracleItems.CatalogNumber"
    Alias ="ViewOrder"
    Expression ="IIf(Left([ItemNumber],2)=Left(Forms!frmMainMenu!NewItemNumber,2),-1,0)"
    Expression ="qryPremiseBase.Base"
    Expression ="qryPremiseBase.DesignCode"
    Alias ="Ext"
    Expression ="Mid([ItemNumber],8,5)"
    Expression ="tblOracleItems.Description"
End
Begin Joins
    LeftTable ="tblOracleItems"
    RightTable ="qryPremiseBase"
    Expression ="tblOracleItems.ItemNumber = qryPremiseBase.OracleItemNumber"
    Flag =1
End
Begin OrderBy
    Expression ="IIf(Left([ItemNumber],2)=Left(Forms!frmMainMenu!NewItemNumber,2),-1,0)"
    Flag =0
    Expression ="tblOracleItems.ItemNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x940836c16577d74f85d41e43d447b4b5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbMemo "Caption" ="Oracle Item Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ext"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbMemo "Caption" ="Oracle Item Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x008064cbbd31ad4f802902b76305c26a
        End
    End
    Begin
        dbText "Name" ="ViewOrder"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d473e7072a6c54e952512034db2c63f
        End
    End
    Begin
        dbText "Name" ="tblOracleItems.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseBase.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOracleItems.CatalogNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-407
    Top =218
    Right =702
    Bottom =740
    Left =-1
    Top =-1
    Right =1077
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =273
        Bottom =149
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
    Begin
        Left =830
        Top =7
        Right =1123
        Bottom =178
        Top =0
        Name ="qryPremiseBase"
        Name =""
    End
End
