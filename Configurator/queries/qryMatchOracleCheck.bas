Operation =1
Option =0
Begin InputTables
    Name ="qryItemNoSearch"
    Name ="qryMatchOracleNumber"
End
Begin OutputColumns
    Expression ="qryMatchOracleNumber.Oracle"
    Alias ="Exists"
    Expression ="IIf(qryMatchOracleNumber![Item No] Is Not Null,-1,0)"
End
Begin Joins
    LeftTable ="qryItemNoSearch"
    RightTable ="qryMatchOracleNumber"
    Expression ="qryItemNoSearch.Oracle = qryMatchOracleNumber.Oracle"
    Flag =2
End
Begin OrderBy
    Expression ="qryMatchOracleNumber.Oracle"
    Flag =1
End
Begin Groups
    Expression ="qryMatchOracleNumber.Oracle"
    GroupLevel =0
    Expression ="IIf(qryMatchOracleNumber![Item No] Is Not Null,-1,0)"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xca1992c24bc9a749ba367258329435b9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Exists"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe8ebb1c6e5def54f928ede6b00a9fe0f
        End
    End
End
Begin
    State =0
    Left =88
    Top =303
    Right =1380
    Bottom =831
    Left =-1
    Top =-1
    Right =1260
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =239
        Bottom =169
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
    Begin
        Left =372
        Top =8
        Right =621
        Bottom =171
        Top =0
        Name ="qryMatchOracleNumber"
        Name =""
    End
End
